import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/pulsecx-backend";
await fs.mkdir(outputDir, { recursive: true });

const wb = Workbook.create();
const dashboard = wb.worksheets.add("Tablero");
const responses = wb.worksheets.add("Respuestas");
const devices = wb.worksheets.add("Dispositivos");
const config = wb.worksheets.add("Configuracion");
const dictionary = wb.worksheets.add("Diccionario");

const navy = "#17335B";
const blue = "#2C5E91";
const paleBlue = "#EAF1F8";
const green = "#26A269";
const yellow = "#F2C94C";
const red = "#D64545";
const gray = "#68768A";
const light = "#F6F8FB";
const border = "#D8E0EA";

function title(sheet, range, text) {
  range.merge();
  range.values = [[text]];
  range.format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF", size: 18 },
    verticalAlignment: "center",
    horizontalAlignment: "left"
  };
  range.format.rowHeight = 34;
}

function header(range) {
  range.format = {
    fill: blue,
    font: { bold: true, color: "#FFFFFF" },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: border }
  };
  range.format.rowHeight = 30;
}

function note(range) {
  range.format = {
    fill: paleBlue,
    font: { color: navy, italic: true },
    wrapText: true,
    verticalAlignment: "center"
  };
}

for (const sheet of [dashboard, responses, devices, config, dictionary]) {
  sheet.showGridLines = false;
}

// Tablero básico: se actualiza con los datos de Respuestas.
title(dashboard, dashboard.getRange("A1:H2"), "PulseCX · Control operativo UCEMA");
dashboard.getRange("A3:H3").merge();
dashboard.getRange("A3:H3").values = [["Resumen directo de la hoja Respuestas. Looker Studio podrá conectarse después a la misma fuente."]];
note(dashboard.getRange("A3:H3"));

dashboard.getRange("A5:B5").merge();
dashboard.getRange("C5:D5").merge();
dashboard.getRange("E5:F5").merge();
dashboard.getRange("G5:H5").merge();
dashboard.getRange("A5:H5").values = [["RESPUESTAS", null, "POSITIVAS", null, "CAPTURAS OK", null, "ESTIMACIONES", null]];
dashboard.getRange("A5:H5").format = { fill: "#DDE8F3", font: { bold: true, color: navy }, horizontalAlignment: "center" };
dashboard.getRange("A6:B7").merge();
dashboard.getRange("C6:D7").merge();
dashboard.getRange("E6:F7").merge();
dashboard.getRange("G6:H7").merge();
dashboard.getRange("A6").formulas = [["=COUNTA('Respuestas'!$B$2:$B$5001)"]];
dashboard.getRange("C6").formulas = [["=IF(A6=0,0,COUNTIF('Respuestas'!$C$2:$C$5001,\"positive\")/A6)"]];
dashboard.getRange("E6").formulas = [["=IF(A6=0,0,COUNTIF('Respuestas'!$H$2:$H$5001,\"captured\")/A6)"]];
dashboard.getRange("G6").formulas = [["=IF(A6=0,0,COUNTIF('Respuestas'!$L$2:$L$5001,\"estimated\")/A6)"]];
dashboard.getRange("A6:H7").format = {
  fill: "#FFFFFF",
  font: { bold: true, color: navy, size: 22 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: border }
};
dashboard.getRange("A6:B7").format.numberFormat = "#,##0";
dashboard.getRange("C6:H7").format.numberFormat = "0.0%";

dashboard.getRange("A10:C10").values = [["Respuesta", "Cantidad", "% del total"]];
header(dashboard.getRange("A10:C10"));
dashboard.getRange("A11:A13").values = [["Muy mala"], ["Neutral"], ["Muy buena"]];
dashboard.getRange("B11:B13").formulas = [
  ["=COUNTIF('Respuestas'!$C$2:$C$5001,\"negative\")"],
  ["=COUNTIF('Respuestas'!$C$2:$C$5001,\"neutral\")"],
  ["=COUNTIF('Respuestas'!$C$2:$C$5001,\"positive\")"]
];
dashboard.getRange("C11").formulas = [["=IF($A$6=0,0,B11/$A$6)"]];
dashboard.getRange("C11:C13").fillDown();
dashboard.getRange("B11:B13").format.numberFormat = "#,##0";
dashboard.getRange("C11:C13").format.numberFormat = "0.0%";
dashboard.getRange("A11:C13").format.borders = { preset: "inside", style: "thin", color: border };
dashboard.getRange("A11:C11").format.fill = "#FBE7E7";
dashboard.getRange("A12:C12").format.fill = "#FFF6D8";
dashboard.getRange("A13:C13").format.fill = "#E5F5EC";

dashboard.getRange("E10:G10").values = [["Rango etario estimado", "Cantidad", "% estimado"]];
header(dashboard.getRange("E10:G10"));
dashboard.getRange("E11:E15").values = [["Menor de 18"], ["18–29"], ["30–49"], ["50+"], ["Incierto"]];
dashboard.getRange("F11:F15").formulas = [
  ["=COUNTIF('Respuestas'!$P$2:$P$5001,\"under_18\")"],
  ["=COUNTIF('Respuestas'!$P$2:$P$5001,\"18_29\")"],
  ["=COUNTIF('Respuestas'!$P$2:$P$5001,\"30_49\")"],
  ["=COUNTIF('Respuestas'!$P$2:$P$5001,\"50_plus\")"],
  ["=COUNTIF('Respuestas'!$P$2:$P$5001,\"uncertain\")"]
];
dashboard.getRange("G11").formulas = [["=IF($G$6=0,0,F11/COUNTIF('Respuestas'!$L$2:$L$5001,\"estimated\"))"]];
dashboard.getRange("G11:G15").fillDown();
dashboard.getRange("F11:F15").format.numberFormat = "#,##0";
dashboard.getRange("G11:G15").format.numberFormat = "0.0%";
dashboard.getRange("E11:G15").format.borders = { preset: "inside", style: "thin", color: border };

dashboard.getRange("A17:C17").values = [["Presentación aparente", "Cantidad", "% estimado"]];
header(dashboard.getRange("A17:C17"));
dashboard.getRange("A18:A20").values = [["Femenina"], ["Masculina"], ["Incierta"]];
dashboard.getRange("B18:B20").formulas = [
  ["=COUNTIF('Respuestas'!$Q$2:$Q$5001,\"feminine\")"],
  ["=COUNTIF('Respuestas'!$Q$2:$Q$5001,\"masculine\")"],
  ["=COUNTIF('Respuestas'!$Q$2:$Q$5001,\"uncertain\")"]
];
dashboard.getRange("C18").formulas = [["=IF($G$6=0,0,B18/COUNTIF('Respuestas'!$L$2:$L$5001,\"estimated\"))"]];
dashboard.getRange("C18:C20").fillDown();
dashboard.getRange("B18:B20").format.numberFormat = "#,##0";
dashboard.getRange("C18:C20").format.numberFormat = "0.0%";
dashboard.getRange("A18:C20").format.borders = { preset: "inside", style: "thin", color: border };

dashboard.getRange("E17:H20").merge();
dashboard.getRange("E17:H20").values = [["Privacidad\n\nLas imágenes no deben almacenarse ni enviarse. Edad y presentación son estimaciones experimentales y sólo deben analizarse de forma agregada."]];
dashboard.getRange("E17:H20").format = { fill: "#FFF3CD", font: { color: "#664D03" }, wrapText: true, verticalAlignment: "center" };
dashboard.freezePanes.freezeRows(3);
dashboard.getRange("A1:H20").format.font.name = "Aptos";
dashboard.getRange("A1:H20").format.columnWidth = 16;
dashboard.getRange("A1:A20").format.columnWidth = 22;
dashboard.getRange("E1:E20").format.columnWidth = 24;

// Hoja principal que recibirá appendRow desde Apps Script.
const responseHeaders = [
  "server_timestamp", "event_id", "rating", "device_id", "location_id", "client_timestamp",
  "app_version", "camera_status", "capture_ms", "capture_width", "capture_height", "model_status",
  "face_detected", "face_count", "face_detection_confidence", "estimated_age_band",
  "apparent_presentation", "presentation_confidence", "inference_ms", "vibration_available",
  "vibration_accepted", "payload_version"
];
responses.getRange("A1:V1").values = [responseHeaders];
header(responses.getRange("A1:V1"));
responses.freezePanes.freezeRows(1);
responses.freezePanes.freezeColumns(2);
responses.getRange("A2:V5001").format.font.name = "Aptos";
responses.getRange("A2:A5001").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
responses.getRange("F2:F5001").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
responses.getRange("I2:K5001").format.numberFormat = "#,##0";
responses.getRange("N2:N5001").format.numberFormat = "#,##0";
responses.getRange("O2:O5001").format.numberFormat = "0.000";
responses.getRange("R2:R5001").format.numberFormat = "0.000";
responses.getRange("S2:S5001").format.numberFormat = "#,##0";
responses.getRange("C2:C5001").dataValidation = { rule: { type: "list", values: ["negative", "neutral", "positive", "diagnostic"] } };
responses.getRange("H2:H5001").dataValidation = { rule: { type: "list", values: ["captured", "not_configured", "permission_denied", "error"] } };
responses.getRange("P2:P5001").dataValidation = { rule: { type: "list", values: ["under_18", "18_29", "30_49", "50_plus", "uncertain", "not_available"] } };
responses.getRange("Q2:Q5001").dataValidation = { rule: { type: "list", values: ["feminine", "masculine", "uncertain", "not_available"] } };
responses.getRange("A1:V5001").format.columnWidth = 16;
responses.getRange("B1:B5001").format.columnWidth = 34;
responses.getRange("D1:F5001").format.columnWidth = 23;
responses.getRange("G1:H5001").format.columnWidth = 20;
responses.getRange("L1:L5001").format.columnWidth = 21;
responses.getRange("O1:R5001").format.columnWidth = 24;

// Registro de dispositivos.
title(devices, devices.getRange("A1:G2"), "Dispositivos PulseCX");
devices.getRange("A3:G3").values = [["device_id", "institucion", "location_id", "descripcion", "activo", "fecha_alta", "notas"]];
header(devices.getRange("A3:G3"));
devices.getRange("A4:G4").values = [["ucema-central-01", "UCEMA", "ucema-sede-central", "Tablet principal · acceso a definir", true, new Date("2026-08-25T00:00:00"), "Editar descripción y ubicación antes de publicar"]];
devices.getRange("E4:E200").dataValidation = { rule: { type: "list", values: [true, false] } };
devices.getRange("F4:F200").format.numberFormat = "yyyy-mm-dd";
devices.freezePanes.freezeRows(3);
devices.getRange("A1:G200").format.font.name = "Aptos";
devices.getRange("A1:G200").format.columnWidth = 20;
devices.getRange("D1:D200").format.columnWidth = 34;
devices.getRange("G1:G200").format.columnWidth = 42;

// Configuración legible por el futuro Apps Script.
title(config, config.getRange("A1:D2"), "Configuración del backend");
config.getRange("A3:D3").values = [["clave", "valor", "editable", "descripcion"]];
header(config.getRange("A3:D3"));
config.getRange("A4:D10").values = [
  ["payload_version", "1.0", false, "Versión del esquema recibido desde la tablet"],
  ["accept_diagnostic_events", false, true, "Guardar o descartar pruebas del panel administrador"],
  ["deduplicate_event_id", true, true, "Evita registrar dos veces el mismo event_id"],
  ["max_payload_bytes", 12000, true, "Tamaño máximo aceptado por el endpoint"],
  ["allowed_device_id", "ucema-central-01", true, "Dispositivo habilitado inicialmente"],
  ["retention_months", 24, true, "Referencia de retención para datos derivados"],
  ["timezone", "America/Argentina/Buenos_Aires", false, "Zona horaria operativa"]
];
config.getRange("C4:C200").dataValidation = { rule: { type: "list", values: [true, false] } };
config.freezePanes.freezeRows(3);
config.getRange("A1:D200").format.font.name = "Aptos";
config.getRange("A1:D200").format.columnWidth = 24;
config.getRange("D1:D200").format.columnWidth = 55;

// Contrato de datos para desarrollar y auditar Apps Script / Looker.
title(dictionary, dictionary.getRange("A1:F2"), "Diccionario de datos · Respuestas");
dictionary.getRange("A3:F3").values = [["campo", "tipo", "requerido", "origen", "valores / formato", "descripcion"]];
header(dictionary.getRange("A3:F3"));
const rows = [
  ["server_timestamp", "datetime", true, "Apps Script", "ISO / fecha-hora", "Hora confiable asignada por el servidor"],
  ["event_id", "text", true, "Tablet", "UUID", "Identificador utilizado para deduplicar"],
  ["rating", "category", true, "Tablet", "negative | neutral | positive", "Respuesta seleccionada"],
  ["device_id", "text", true, "Tablet", "ID registrado", "Identifica la tablet"],
  ["location_id", "text", true, "Tablet", "ID de sede", "Ubicación operativa"],
  ["client_timestamp", "datetime", true, "Tablet", "ISO", "Hora informada por el dispositivo"],
  ["app_version", "text", true, "Tablet", "web-prototype-x.y.z", "Versión del frontend"],
  ["camera_status", "category", true, "Tablet", "captured | not_configured | permission_denied | error", "Resultado de la cámara"],
  ["capture_ms", "integer", false, "Tablet", "milisegundos", "Tiempo hasta capturar el fotograma"],
  ["capture_width", "integer", false, "Tablet", "píxeles", "Ancho de la captura"],
  ["capture_height", "integer", false, "Tablet", "píxeles", "Alto de la captura"],
  ["model_status", "category", true, "Tablet", "estimated | no_face | multiple_faces | error | not_run", "Resultado del modelo local"],
  ["face_detected", "boolean", true, "Tablet", "TRUE | FALSE", "Indica detección de al menos un rostro"],
  ["face_count", "integer", true, "Tablet", "0 o más", "Cantidad de rostros detectados"],
  ["face_detection_confidence", "decimal", false, "Tablet", "0–1", "Confianza del detector cuando existe un único rostro"],
  ["estimated_age_band", "category", false, "Tablet", "under_18 | 18_29 | 30_49 | 50_plus | uncertain", "Rango etario estimado; no es edad objetiva"],
  ["apparent_presentation", "category", false, "Tablet", "feminine | masculine | uncertain", "Presentación aparente estimada; no es identidad de género"],
  ["presentation_confidence", "decimal", false, "Tablet", "0–1", "Confianza de la clasificación aparente"],
  ["inference_ms", "integer", false, "Tablet", "milisegundos", "Duración del análisis local"],
  ["vibration_available", "boolean", true, "Tablet", "TRUE | FALSE", "Disponibilidad reportada por el navegador"],
  ["vibration_accepted", "boolean", true, "Tablet", "TRUE | FALSE", "Resultado de la solicitud de vibración"],
  ["payload_version", "text", true, "Apps Script", "1.0", "Versión del contrato almacenado"]
];
dictionary.getRange(`A4:F${rows.length + 3}`).values = rows;
dictionary.getRange(`C4:C${rows.length + 3}`).dataValidation = { rule: { type: "list", values: [true, false] } };
dictionary.freezePanes.freezeRows(3);
dictionary.getRange(`A1:F${rows.length + 3}`).format.font.name = "Aptos";
dictionary.getRange(`A1:F${rows.length + 3}`).format.wrapText = true;
dictionary.getRange(`A1:F${rows.length + 3}`).format.columnWidth = 22;
dictionary.getRange(`E1:E${rows.length + 3}`).format.columnWidth = 48;
dictionary.getRange(`F1:F${rows.length + 3}`).format.columnWidth = 52;
dictionary.getRange(`A4:F${rows.length + 3}`).format.borders = { preset: "inside", style: "thin", color: border };

const dashCheck = await wb.inspect({ kind: "table", sheetId: "Tablero", range: "A1:H20", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 8 });
console.log(dashCheck.ndjson);
const errorCheck = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errorCheck.ndjson);

for (const [sheetName, range] of [
  ["Tablero", "A1:H20"],
  ["Respuestas", "A1:V8"],
  ["Dispositivos", "A1:G8"],
  ["Configuracion", "A1:D12"],
  ["Diccionario", `A1:F${rows.length + 3}`]
]) {
  const preview = await wb.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/preview-${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(`${outputDir}/PulseCX_Backend_UCEMA.xlsx`);
console.log(`SAVED ${outputDir}/PulseCX_Backend_UCEMA.xlsx`);
