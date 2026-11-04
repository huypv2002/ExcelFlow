/**
 * ExcelFlow Suite — Trang chủ Tối giản (Chỉ có các nút bấm chức năng)
 */
(function() {
  'use strict';

  function showToast(message) {
    if (typeof window.efToast === 'function') {
      window.efToast(message);
    } else {
      const t = document.getElementById('toast');
      if (t) {
        t.textContent = message;
        t.classList.add('show');
        clearTimeout(showToast.timer);
        showToast.timer = setTimeout(() => t.classList.remove('show'), 2500);
      }
    }
  }

  function init() {
    // Gắn sự kiện click cho các nút bấm trên Trang chủ
    const cards = document.querySelectorAll('.ef-btn-card');

    cards.forEach(card => {
      card.addEventListener('click', (e) => {
        e.preventDefault();
        const appId = card.dataset.appId;
        const appTitle = card.dataset.appTitle || '';

        if (appId === 'excelflow') {
          // Mở công cụ ExcelFlow
          if (typeof window.efGoPage === 'function') {
            window.efGoPage('import');
          } else {
            const navImport = document.getElementById('navImport');
            if (navImport) navImport.click();
          }
        } else {
          // Thông báo Toast nhẹ nhàng, không mở popup
          showToast(`⚡ Chức năng "${appTitle}" đang được phát triển!`);
        }
      });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
