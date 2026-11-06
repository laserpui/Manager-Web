// Keep the appearance of standalone pages and pages embedded in Manager Web in sync.
(() => {
  const key = 'manager-web-theme';
  const apply = (theme, persist = false, notifyParent = false) => {
    const next = theme === 'dark' ? 'dark' : 'light';
    document.documentElement.dataset.theme = next;
    document.documentElement.style.colorScheme = next;
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.content = next === 'dark' ? '#0b1422' : '#e5ffff';
    if (document.documentElement.classList.contains('ui-stock') && typeof window.applyTheme === 'function') {
      window.applyTheme(next, false);
    }
    if (persist) {
      try { localStorage.setItem(key, next); } catch (_) {}
    }
    if (notifyParent && window.parent !== window) {
      window.parent.postMessage({ type: 'manager-web-theme', theme: next }, location.origin);
    }
  };
  window.ManagerTheme = { apply };
  try {
    const saved = localStorage.getItem(key);
    if (saved === 'dark' || saved === 'light') apply(saved);
  } catch (_) {}
  window.addEventListener('message', event => {
    if (event.origin !== location.origin || event.source !== window.parent || event.data?.type !== 'manager-web-theme') return;
    apply(event.data.theme, true);
  });
  window.addEventListener('storage', event => {
    if (event.key === key && (event.newValue === 'dark' || event.newValue === 'light')) apply(event.newValue);
  });
})();
