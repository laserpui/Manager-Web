// Read-only presentation for pages opened from the Manager Web index.
// The query flag keeps each original standalone page available for its own workflow.
(() => {
  if (new URLSearchParams(location.search).get('view') !== '1') return;

  document.documentElement.dataset.viewOnly = 'true';
  const style = document.createElement('style');
  style.textContent = `
    html[data-view-only] a[href="admin.html"],
    html[data-view-only] a[href$="/admin.html"],
    html[data-view-only] .btn-delete-row,
    html[data-view-only] [data-action="open-form"],
    html[data-view-only] [data-return-id],
    html[data-view-only] #transactionPage,
    html[data-view-only] #returnModal { display: none !important; }
  `;

  const page = decodeURIComponent(location.pathname);
  if (page.includes('รายงานทำงานวันหยุด')) {
    style.textContent += `
      html[data-view-only] .dashboard-grid { grid-template-columns: minmax(0, 1fr) !important; }
      html[data-view-only] .dashboard-grid > section:first-child,
      html[data-view-only] #recordsTable tr > :last-child { display: none !important; }
    `;
  }
  document.head.append(style);

  const originalFetch = window.fetch;
  window.fetch = function (input, init) {
    const method = String(init?.method || (input instanceof Request ? input.method : 'GET')).toUpperCase();
    if (method !== 'GET' && method !== 'HEAD') {
      return Promise.reject(new Error('หน้านี้เปิดเพื่อดูข้อมูลเท่านั้น'));
    }
    return originalFetch.apply(this, arguments);
  };

  const originalOpen = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (method) {
    if (!['GET', 'HEAD'].includes(String(method).toUpperCase())) {
      throw new Error('หน้านี้เปิดเพื่อดูข้อมูลเท่านั้น');
    }
    return originalOpen.apply(this, arguments);
  };

  navigator.sendBeacon = () => false;
  document.addEventListener('submit', event => {
    if (event.target.matches('#reportForm, #borrowForm')) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  }, true);
})();
