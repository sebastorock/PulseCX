# PulseCX UCEMA — prototipo web para tablet

Este repositorio contiene la PWA que se instala en la tablet. La versión actual incluye:

- interfaz de tres respuestas;
- pantalla completa y orientación horizontal;
- animación y solicitud de vibración en `pointerdown`;
- apertura de la cámara frontal en cada voto;
- captura del primer fotograma únicamente en memoria;
- apagado inmediato de la cámara;
- detección local de un único rostro;
- estimación experimental de rango etario y presentación aparente;
- descarte automático del resultado cuando no hay un rostro o aparecen varios;
- medición de tiempo y resolución de captura;
- funcionamiento offline de la interfaz mediante service worker;
- envío simulado, todavía sin backend.

No almacena ni transmite fotografías. El evento preparado para el futuro HTTP POST contiene sólo datos derivados, nunca la imagen.

## Publicación en GitHub Pages

Subir todos los archivos y conservar la estructura de carpetas. Publicar desde la raíz del repositorio mediante GitHub Pages y abrir la URL HTTPS resultante.

Después de actualizar una versión ya instalada, cerrar y volver a abrir la PWA. Si persiste una versión anterior, borrar los datos del sitio o desinstalar y reinstalar el acceso PWA.

## Preparación de la tablet

1. Abrir `https://URL-DE-PULSECX/?admin=1`.
2. Pulsar **Probar cámara** y aceptar el permiso.
3. Pulsar **Probar vibración**.
4. Confirmar que la captura informe tiempo, resolución y una estimación (o el motivo por el que no pudo obtenerla).
5. Cerrar el panel y abrir la URL normal.
6. Instalar la PWA desde Chrome o Samsung Internet.

La Vibration API puede informar que la solicitud fue aceptada aunque el equipo no tenga actuador háptico. En ese caso el feedback operativo es la animación visual.

## Prueba de cámara

Cada presión sobre una carita solicita la cámara frontal a 640×480 y 10 FPS, captura el primer fotograma y detiene todas las pistas de cámara. Después analiza la imagen localmente, prepara los campos derivados y limpia la memoria. El resultado técnico se imprime en la consola y queda visible en el panel administrativo durante esa sesión.

Los valores demográficos son inferencias aproximadas, no datos objetivos. La salida utiliza rangos amplios (`under_18`, `18_29`, `30_49`, `50_plus`) y presentación aparente (`masculine`, `feminine`, `uncertain`). No debe usarse para identificar personas ni tomar decisiones individuales.

## Próximas etapas

1. Cola local de eventos con IndexedDB.
2. Google Sheets y Apps Script.
3. HTTP POST, deduplicación y reintentos.
4. Dashboard en Looker Studio.
