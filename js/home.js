/**
 * ExcelFlow Suite — App Launcher Manager (Chỉ xử lý các nút bấm)
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
        showToast.timer = setTimeout(() => t.classList.remove('show'), 3000);
      }
    }
  }

  function init() {
    // Gắn sự kiện click cho tất cả các nút bấm chức năng trên Trang chủ
    const appButtons = document.querySelectorAll('.ef-app-btn');

    appButtons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const appId = btn.dataset.appId;
        const appTitle = btn.dataset.appTitle || 'này';

        if (appId === 'excelflow') {
          // Mở ngay công cụ ExcelFlow
          if (typeof window.efGoPage === 'function') {
            window.efGoPage('import');
          } else {
            const navImport = document.getElementById('navImport');
            if (navImport) navImport.click();
          }
        } else {
          // Thông báo Toast nhẹ nhàng 2.5 giây, không mở bất kỳ popup nào
          showToast(`⚡ Tính năng "${appTitle}" đang được hoàn thiện! Vui lòng dùng tính năng "Ghép & Đối chiếu Excel".`);
        }
      });
    });
  }

  // Khởi động khi DOM sẵn sàng
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
