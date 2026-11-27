self.addEventListener('install', function (event) {
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('push', function (event) {
  var data = {};
  if (event.data) {
    try {
      data = event.data.json();
    } catch (e) {
      try {
        data = { body: event.data.text() };
      } catch (e2) {
        data = {};
      }
    }
  }

  var title = data.title || 'SysManager — Lembrete Diário';
  var options = {
    body: data.body || 'Você tem atualizações na sua esteira do SysManager.',
    icon: data.icon || '/icon-192.png',
    badge: data.badge || '/badge-96.png',
    tag: data.tag || 'sysmanager-daily',
    renotify: data.renotify !== false,
    requireInteraction: data.requireInteraction !== false,
    vibrate: [150, 80, 150],
    data: {
      url: data.url || '/?briefing=1',
      tipo: data.tipo || 'briefing'
    }
  };

  var isBriefing = options.data.tipo === 'briefing' || (options.tag || '').indexOf('sysmanager-daily-') === 0;
  var tarefas = [self.registration.showNotification(title, options)];
  if (isBriefing && typeof caches !== 'undefined') {
    tarefas.push(
      caches.open('sysmanager-push-state').then(function (cache) {
        return cache.put('/__last_briefing_tag', new Response(String(options.tag || '')));
      }).catch(function () {})
    );
  }

  event.waitUntil(Promise.all(tarefas));
});

self.addEventListener('notificationclick', function (event) {
  event.notification.close();

  var rawUrl = (event.notification && event.notification.data && event.notification.data.url) || '/?briefing=1';
  var targetUrl;
  try {
    targetUrl = new URL(rawUrl, self.location.origin).href;
  } catch (e) {
    targetUrl = self.location.origin + '/?briefing=1';
  }

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (clientList) {
      for (var i = 0; i < clientList.length; i++) {
        var client = clientList[i];
        if ((client.url || '').indexOf(self.location.origin) === 0 && 'focus' in client) {
          return client.focus().then(function (focusedClient) {
            try {
              (focusedClient || client).postMessage({
                type: 'OPEN_DAILY_BRIEFING',
                url: targetUrl
              });
            } catch (errMsg) {}
            return focusedClient;
          });
        }
      }
      if (self.clients.openWindow) {
        return self.clients.openWindow(targetUrl);
      }
      return null;
    })
  );
});
