const RESPONSE_HEADERS = Object.freeze([
  'server_timestamp', 'event_id', 'rating', 'device_id', 'location_id',
  'client_timestamp', 'app_version', 'camera_status', 'capture_ms',
  'capture_width', 'capture_height', 'model_status', 'face_detected',
  'face_count', 'face_detection_confidence', 'estimated_age_band',
  'apparent_presentation', 'presentation_confidence', 'inference_ms',
  'vibration_available', 'vibration_accepted', 'payload_version'
]);

const ALLOWED_RATINGS = ['negative', 'neutral', 'positive', 'diagnostic'];
const ALLOWED_CAMERA_STATUSES = [
  'captured', 'not_configured', 'permission_denied', 'error', 'starting'
];
const ALLOWED_MODEL_STATUSES = [
  'estimated', 'no_face', 'multiple_faces', 'error', 'not_run', 'pending'
];

/** Permite comprobar en el navegador que el servicio está publicado. */
function doGet() {
  return jsonResponse_({
    ok: true,
    service: 'PulseCX Backend UCEMA',
    version: '1.0',
    timestamp: new Date().toISOString()
  });
}

/** Recibe un voto enviado por la tablet y lo agrega a Respuestas. */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const config = getConfig_();
    const rawBody = e && e.postData && e.postData.contents
      ? e.postData.contents
      : '';
    const maxBytes = Number(config.max_payload_bytes || 12000);

    if (!rawBody) {
      return jsonResponse_({ ok: false, error: 'empty_body' });
    }

    if (Utilities.newBlob(rawBody).getBytes().length > maxBytes) {
      return jsonResponse_({ ok: false, error: 'payload_too_large' });
    }

    let payload;
    try {
      payload = JSON.parse(rawBody);
    } catch (error) {
      return jsonResponse_({ ok: false, error: 'invalid_json' });
    }

    const validationError = validatePayload_(payload, config);
    if (validationError) {
      return jsonResponse_({ ok: false, error: validationError });
    }

    if (
      payload.rating === 'diagnostic' &&
      !asBoolean_(config.accept_diagnostic_events)
    ) {
      return jsonResponse_({
        ok: true,
        accepted: false,
        reason: 'diagnostic_ignored'
      });
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Respuestas');
    if (!sheet) throw new Error('No existe la hoja Respuestas');

    validateHeaders_(sheet);

    if (
      asBoolean_(config.deduplicate_event_id) &&
      eventExists_(sheet, payload.event_id)
    ) {
      return jsonResponse_({
        ok: true,
        accepted: false,
        duplicate: true,
        event_id: payload.event_id
      });
    }

    const serverTimestamp = new Date();
    const row = [
      serverTimestamp,
      cleanText_(payload.event_id, 80),
      cleanText_(payload.rating, 20),
      cleanText_(payload.device_id, 80),
      cleanText_(payload.location_id, 120),
      parseDateOrBlank_(payload.client_timestamp),
      cleanText_(payload.app_version, 40),
      cleanText_(payload.camera_status, 30),
      numberOrBlank_(payload.capture_ms),
      numberOrBlank_(payload.capture_width),
      numberOrBlank_(payload.capture_height),
      cleanText_(payload.model_status, 30),
      booleanOrBlank_(payload.face_detected),
      numberOrBlank_(payload.face_count),
      probabilityOrBlank_(payload.face_detection_confidence),
      cleanText_(payload.estimated_age_band, 30),
      cleanText_(payload.apparent_presentation, 30),
      probabilityOrBlank_(payload.presentation_confidence),
      numberOrBlank_(payload.inference_ms),
      booleanOrBlank_(payload.vibration_available),
      booleanOrBlank_(payload.vibration_accepted),
      cleanText_(config.payload_version || '1.0', 20)
    ];

    sheet.appendRow(row);

    return jsonResponse_({
      ok: true,
      accepted: true,
      event_id: payload.event_id,
      server_timestamp: serverTimestamp.toISOString()
    });
  } catch (error) {
    console.error(error);
    return jsonResponse_({
      ok: false,
      error: 'server_error',
      detail: String(error.message || error)
    });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function validatePayload_(payload, config) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return 'invalid_payload';
  }
  if (!payload.event_id || typeof payload.event_id !== 'string') {
    return 'missing_event_id';
  }
  if (ALLOWED_RATINGS.indexOf(payload.rating) === -1) {
    return 'invalid_rating';
  }
  if (!payload.device_id || typeof payload.device_id !== 'string') {
    return 'missing_device_id';
  }
  if (!payload.location_id || typeof payload.location_id !== 'string') {
    return 'missing_location_id';
  }
  if (
    !payload.client_timestamp ||
    isNaN(new Date(payload.client_timestamp).getTime())
  ) {
    return 'invalid_client_timestamp';
  }
  if (ALLOWED_CAMERA_STATUSES.indexOf(payload.camera_status) === -1) {
    return 'invalid_camera_status';
  }
  if (ALLOWED_MODEL_STATUSES.indexOf(payload.model_status) === -1) {
    return 'invalid_model_status';
  }

  const allowedDevice = String(config.allowed_device_id || '').trim();
  if (allowedDevice && payload.device_id !== allowedDevice) {
    return 'device_not_allowed';
  }

  return null;
}

function validateHeaders_(sheet) {
  const actual = sheet
    .getRange(1, 1, 1, RESPONSE_HEADERS.length)
    .getDisplayValues()[0];
  const valid = RESPONSE_HEADERS.every(function(header, index) {
    return actual[index] === header;
  });

  if (!valid) {
    throw new Error('Las columnas de Respuestas fueron modificadas');
  }
}

function eventExists_(sheet, eventId) {
  if (sheet.getLastRow() < 2) return false;

  const match = sheet
    .getRange(2, 2, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(eventId))
    .matchEntireCell(true)
    .findNext();

  return Boolean(match);
}

function getConfig_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Configuracion');

  if (!sheet) throw new Error('No existe la hoja Configuracion');
  if (sheet.getLastRow() < 4) return {};

  const values = sheet
    .getRange(4, 1, sheet.getLastRow() - 3, 2)
    .getValues();

  return values.reduce(function(result, row) {
    const key = String(row[0] || '').trim();
    if (key) result[key] = row[1];
    return result;
  }, {});
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function cleanText_(value, maxLength) {
  if (value === null || value === undefined) return '';
  return String(value).trim().slice(0, maxLength);
}

function numberOrBlank_(value) {
  if (value === null || value === undefined || value === '') return '';
  const number = Number(value);
  return Number.isFinite(number) ? number : '';
}

function probabilityOrBlank_(value) {
  const number = numberOrBlank_(value);
  if (number === '') return '';
  return Math.min(1, Math.max(0, number));
}

function booleanOrBlank_(value) {
  if (value === true || value === false) return value;
  return '';
}

function parseDateOrBlank_(value) {
  const date = new Date(value);
  return isNaN(date.getTime()) ? '' : date;
}

function asBoolean_(value) {
  return value === true ||
    String(value).toLowerCase() === 'true' ||
    Number(value) === 1;
}
