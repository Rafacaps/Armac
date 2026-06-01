// Sem cache — sempre busca da rede
self.addEventListener('install', function(e){
  self.skipWaiting();
});

self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.map(function(key){ return caches.delete(key); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e){
  // Não interceptar nada — sempre vai para a rede
  if(e.request.url.includes('supabase.co')) return;
  e.respondWith(fetch(e.request).catch(function(){
    return caches.match(e.request);
  }));
});
