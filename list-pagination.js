// Display-only pagination shared by the five embedded websites.
window.createListPager = function createListPager(anchor, render) {
  const size = 10;
  let currentPage = 1;
  const bar = document.createElement('nav');
  bar.className = 'list-pagination';
  bar.hidden = true;
  bar.setAttribute('aria-label', 'หน้ารายการ');
  bar.innerHTML = '<span class="list-pagination-summary" aria-live="polite"></span><div class="list-pagination-actions"><button type="button" data-page="previous">← ย้อนกลับ</button><span class="list-pagination-page"></span><button type="button" data-page="next">หน้าถัดไป →</button></div>';
  anchor.insertAdjacentElement('afterend', bar);
  const previous = bar.querySelector('[data-page="previous"]');
  const next = bar.querySelector('[data-page="next"]');
  previous.addEventListener('click', () => { currentPage--; render(); });
  next.addEventListener('click', () => { currentPage++; render(); });
  return {
    get offset() { return (currentPage - 1) * size; },
    reset() { currentPage = 1; },
    page(items) {
      const totalPages = Math.max(1, Math.ceil(items.length / size));
      currentPage = Math.max(1, Math.min(currentPage, totalPages));
      const start = (currentPage - 1) * size;
      const shown = items.slice(start, start + size);
      bar.hidden = items.length === 0;
      bar.querySelector('.list-pagination-summary').textContent = items.length ? `แสดง ${start + 1}–${start + shown.length} จาก ${items.length} รายการ` : 'ไม่มีรายการ';
      bar.querySelector('.list-pagination-page').textContent = `หน้า ${currentPage} / ${totalPages}`;
      previous.disabled = currentPage === 1;
      next.disabled = currentPage === totalPages;
      return shown;
    }
  };
};
