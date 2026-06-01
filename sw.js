const CACHE_NAME = 'armac-taboca-v1';
const ASSETS = [
  '/',
  '/index.html',
  'https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// Instalar — cachear assets principais
self.addEventListener('install', function(e){
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){
      return cache.addAll(ASSETS).catch(function(err){
        console.log('Cache parcial:', err);
      });
    })
  );
  self.skipWaiting();
});

// Ativar — limpar caches antigos
self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(
        keys.filter(function(key){ return key !== CACHE_NAME; })
            .map(function(key){ return caches.delete(key); })
      );
    })
  );
  self.clients.claim();
});

// Fetch — servir do cache quando offline
self.addEventListener('fetch', function(e){
  // Não interceptar requests do Supabase (dados sempre vêm da rede)
  if(e.request.url.includes('supabase.co')){
    return;
  }
  e.respondWith(
    caches.match(e.request).then(function(cached){
      return cached || fetch(e.request).then(function(response){
        // Cachear novos assets
        if(response && response.status === 200 && response.type === 'basic'){
          var clone = response.clone();
          caches.open(CACHE_NAME).then(function(cache){
            cache.put(e.request, clone);
          });
        }
        return response;
      }).catch(function(){
        return cached;
      });
    })
  );
});
