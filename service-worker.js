const CACHE_NAME = "pulsecx-ucema-v8";
const APP_FILES = [
  "./",
  "./index.html",
  "./site.webmanifest",
  "./vendor/face-api.min.js",
  "./vendor/ort.min.js",
  "./vendor/ort-wasm-simd-threaded.wasm",
  "./vendor/ort-wasm-simd-threaded.jsep.mjs",
  "./vendor/ort-wasm-simd-threaded.jsep.wasm",
  "./models/tiny_face_detector_model-weights_manifest.json",
  "./models/tiny_face_detector_model-shard1",
  "./models/age_gender_model-weights_manifest.json",
  "./models/age_gender_model-shard1",
  "./models/face_landmark_68_model-weights_manifest.json",
  "./models/face_landmark_68_model-shard1",
  "./models/face_recognition_model-weights_manifest.json",
  "./models/face_recognition_model-shard1",
  "./models/face_recognition_model-shard2",
  "./models/face_capture_quality.onnx",
  "./assets/icons/pulsecx-192.png",
  "./assets/icons/pulsecx-512.png",
  "./assets/emojis/pulse-negative-red.png",
  "./assets/emojis/pulse-neutral-yellow.png",
  "./assets/emojis/pulse-positive-green.png"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
    ))
  );
  self.clients.claim();
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request).then(response => response || caches.match("./index.html")))
  );
});
