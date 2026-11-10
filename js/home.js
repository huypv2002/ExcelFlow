/**
 * Puddin Excel Flow — Trang chủ Tối giản & Đổi màu giao diện (Xanh lá / Đỏ)
 */
(function() {
  'use strict';

  const THEME_STORAGE_KEY = 'puddin_ef_theme';

  function applyTheme(theme) {
    const validTheme = theme === 'red' ? 'red' : 'green';
    document.documentElement.setAttribute('data-theme', validTheme);
    try {
      localStorage.setItem(THEME_STORAGE_KEY, validTheme);
    } catch (e) {}

    document.querySelectorAll('.ef-theme-btn[data-theme-choice]').forEach(btn => {
      const isSelected = btn.dataset.themeChoice === validTheme;
      btn.classList.toggle('active', isSelected);
      btn.setAttribute('aria-pressed', isSelected ? 'true' : 'false');
    });
  }

  function initThemeSwitcher() {
    let saved = 'green';
    try {
      const stored = localStorage.getItem(THEME_STORAGE_KEY);
      if (stored === 'red' || stored === 'green') saved = stored;
    } catch (e) {}
    applyTheme(saved);

    document.querySelectorAll('.ef-theme-btn[data-theme-choice]').forEach(btn => {
      btn.addEventListener('click', () => {
        applyTheme(btn.dataset.themeChoice);
      });
    });
  }

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
    initThemeSwitcher();

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
        } else if (appId === 'merger') {
          // Mở công cụ Gộp nhiều file
          if (typeof window.efGoPage === 'function') {
            window.efGoPage('merger');
          } else {
            const navMerger = document.getElementById('navMerger');
            if (navMerger) navMerger.click();
          }
        } else if (appId === 'pdf2excel') {
          // Mở công cụ Chuyển từ PDF sang Excel
          if (typeof window.efGoPage === 'function') {
            window.efGoPage('pdf2excel');
          } else {
            const navPdf = document.getElementById('navPdf2Excel');
            if (navPdf) navPdf.click();
          }
        } else {
          // Thông báo Toast nhẹ nhàng, không mở popup
          showToast(`Chức năng "${appTitle}" đang được phát triển!`);
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
