// auth.js — Control de acceso compartido
(function () {
  var KEY = 'pg_session';
  var _origFetch = window.fetch;
  var _scriptUrl = ''; // detectado del primer request al Apps Script

  function getSession() {
    try {
      var raw = localStorage.getItem(KEY);
      if (!raw) return null;
      var s = JSON.parse(raw);
      if (Date.now() > s.exp) { localStorage.removeItem(KEY); return null; }
      return s;
    } catch (e) { localStorage.removeItem(KEY); return null; }
  }

  function showExpiredToast() {
    if (document.getElementById('auth-expired-toast')) return;
    var t = document.createElement('div');
    t.id = 'auth-expired-toast';
    t.style.cssText = 'position:fixed;top:16px;right:16px;z-index:9999;background:#c62828;color:#fff;' +
      'padding:13px 18px;border-radius:8px;font-size:13px;font-weight:600;' +
      'box-shadow:0 4px 16px rgba(0,0,0,.2);max-width:280px;line-height:1.4';
    t.textContent = 'Tu sesión expiró. Volvé a iniciar sesión.';
    (document.body || document.documentElement).appendChild(t);
    setTimeout(function () {
      sessionStorage.setItem('auth_redirect', window.location.pathname.split('/').pop() || 'index.html');
      window.location.replace('login.html');
    }, 2500);
  }

  window.AUTH = {
    checkAuth: function () {
      var s = getSession();
      if (!s) {
        sessionStorage.setItem('auth_redirect', window.location.pathname.split('/').pop() || 'index.html');
        window.location.replace('login.html');
        return null;
      }
      var av = document.getElementById('user-av');
      if (av) {
        av.textContent = s.initials;
        av.title = s.name + ' — ' + s.email;
        av.setAttribute('aria-label', 'Usuario: ' + s.name);
        av.style.cursor = 'pointer';

        // Menú desplegable (reemplaza al confirm())
        var dd = document.createElement('div');
        dd.id = 'auth-dd';
        dd.style.cssText = 'display:none;position:fixed;top:58px;right:12px;z-index:9000;' +
          'background:#fff;border:1px solid #dce4e2;border-radius:10px;' +
          'box-shadow:0 4px 20px rgba(0,0,0,.13);min-width:210px;padding:14px;font-family:inherit;';
        var nm = document.createElement('div');
        nm.style.cssText = 'font-weight:700;font-size:13px;color:#1a1c1b;margin-bottom:2px';
        nm.textContent = s.name;
        var em = document.createElement('div');
        em.style.cssText = 'font-size:11.5px;color:#5f6362;margin-bottom:12px';
        em.textContent = s.email;
        var btn = document.createElement('button');
        btn.style.cssText = 'width:100%;padding:8px;background:#ba1a1a;color:#fff;border:none;' +
          'border-radius:6px;cursor:pointer;font-size:12.5px;font-weight:600;font-family:inherit';
        btn.textContent = 'Cerrar sesión';
        btn.onclick = function () { dd.style.display = 'none'; AUTH.logout(); };
        dd.appendChild(nm);
        dd.appendChild(em);
        dd.appendChild(btn);
        document.body.appendChild(dd);

        av.addEventListener('click', function (e) {
          e.stopPropagation();
          dd.style.display = dd.style.display === 'none' ? 'block' : 'none';
        });
        document.addEventListener('click', function () { dd.style.display = 'none'; });
      }
      return s;
    },

    saveSession: function (name, email, token) {
      var parts = name.trim().split(/\s+/);
      var initials = parts.length >= 2
        ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
        : name.slice(0, 2).toUpperCase();
      var data = { name: name, email: email, initials: initials, token: token || '', exp: Date.now() + 6 * 3600 * 1000 };
      localStorage.setItem(KEY, JSON.stringify(data));
      return data;
    },

    logout: function () {
      var s = getSession();
      if (s && s.token && _scriptUrl) {
        // Invalidar token en el servidor (fire-and-forget)
        try {
          _origFetch(_scriptUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'logout', token: s.token })
          });
        } catch (e) { /* si falla la red, el token expira solo en 6h */ }
      }
      localStorage.removeItem(KEY);
      sessionStorage.clear();
      window.location.replace('login.html');
    },

    getSession: getSession
  };

  // Interceptar fetch: inyectar token + detectar sesión expirada en servidor
  window.fetch = function (url, opts) {
    var u = String(url || '');
    if (u.indexOf('script.google.com/macros/s/') >= 0) {
      if (!_scriptUrl) _scriptUrl = u.split('?')[0];
      var tok = '';
      try { tok = (getSession() || {}).token || ''; } catch (e) {}
      if (tok) {
        if (opts && String(opts.method || '').toUpperCase() === 'POST') {
          try {
            var b = JSON.parse(opts.body || '{}');
            if (!b.token) { b.token = tok; opts = Object.assign({}, opts, { body: JSON.stringify(b) }); }
          } catch (e) {}
        } else {
          var sep = u.indexOf('?') >= 0 ? '&' : '?';
          url = u + sep + 'token=' + encodeURIComponent(tok);
        }
      }
    }
    var p = _origFetch.apply(this, [url, opts]);
    if (u.indexOf('script.google.com/macros/s/') >= 0) {
      p = p.then(function (resp) {
        return resp.clone().text().then(function (text) {
          try { if (JSON.parse(text).code === 401) showExpiredToast(); } catch (e) {}
          return resp;
        }).catch(function () { return resp; });
      });
    }
    return p;
  };

})();
