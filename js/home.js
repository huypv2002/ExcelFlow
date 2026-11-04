/**
 * ExcelFlow Suite — Trang chủ App Workspace Hub & Ecosystem Manager
 */
(function() {
  'use strict';

  // Danh mục công cụ trong bức tranh toàn cảnh hệ sinh thái Excel
  const TOOLS = [
    {
      id: 'excelflow',
      title: 'Ghép & Đối chiếu dữ liệu (ExcelFlow)',
      category: 'matching',
      categoryName: 'Đối chiếu & Nhập liệu',
      status: 'active',
      statusText: 'Sẵn sàng dùng',
      icon: '⚡',
      iconTheme: 'icon-teal',
      featured: true,
      desc: 'Tự động phân tích và đối chiếu dữ liệu từ nhiều file sao kê hoặc bảng tính nguồn vào file mẫu chuẩn. Hỗ trợ ghép theo cột khóa hoặc số dòng, phát hiện lệch số tiền và CRUD dữ liệu trực tiếp.',
      tags: ['Đa file nguồn', 'Sao kê ngân hàng', 'Ghép cột', 'Lưu IndexedDB'],
      actionText: 'Mở công cụ ngay →',
      details: {
        highlights: [
          'Ghép đồng thời nhiều file nguồn (sao kê các tháng) vào 1 file đích chuẩn.',
          'Thuật toán chuyên sâu cho sao kê ngân hàng: đảo ngược thứ tự, bỏ dòng tổng cộng.',
          'Xem trước bảng dữ liệu đã ghép, lọc theo trạng thái (sẽ nhập, đã ghép, bỏ qua).',
          'Popup tìm kiếm, thêm, sửa, xóa dòng dữ liệu trực tiếp trước khi kết xuất file.',
          'Lưu trữ lịch sử an toàn trong trình duyệt (IndexedDB), xuất báo cáo CSV chi tiết.'
        ]
      }
    },
    {
      id: 'merger',
      title: 'Gộp nhiều File & Sheet thành một',
      category: 'processing',
      categoryName: 'Xử lý & Cấu trúc',
      status: 'upcoming',
      statusText: 'Sắp ra mắt • Q2/2026',
      icon: '📚',
      iconTheme: 'icon-indigo',
      featured: false,
      desc: 'Tập hợp hàng chục file Excel báo cáo định kỳ từ các phòng ban, chi nhánh hoặc gộp nhiều Sheet thành 1 bảng tổng hợp duy nhất chỉ trong vài giây. Tự động căn chỉnh theo tiêu đề cột.',
      tags: ['Gộp hàng loạt', 'Tự căn cột', 'Báo cáo tổng hợp', 'Giữ định dạng'],
      actionText: 'Xem chi tiết tính năng',
      details: {
        highlights: [
          'Gộp không giới hạn số lượng file Excel vào 1 bảng tổng hợp duy nhất.',
          'Tự động đối sánh tên cột giống nhau, tạo thêm cột mới cho các dữ liệu đặc thù.',
          'Hỗ trợ thêm cột tự động lưu tên file gốc hoặc tên Sheet để truy xuất nguồn gốc.',
          'Chạy 100% trong RAM trình duyệt, không lo bị gián đoạn khi file dung lượng lớn.'
        ]
      }
    },
    {
      id: 'splitter',
      title: 'Tách File theo Điều kiện thông minh',
      category: 'processing',
      categoryName: 'Xử lý & Cấu trúc',
      status: 'upcoming',
      statusText: 'Sắp ra mắt • Q2/2026',
      icon: '✂️',
      iconTheme: 'icon-orange',
      featured: false,
      desc: 'Tự động phân tách 1 file Excel lớn thành nhiều file nhỏ riêng biệt hoặc nhiều Sheet theo giá trị cột phân loại (Chi nhánh, Tháng, Mã nhân viên, Khách hàng) và xuất file ZIP chuẩn.',
      tags: ['Tách theo cột', 'Phân loại chi nhánh', 'Xuất file ZIP', 'Đặt tên tự động'],
      actionText: 'Xem chi tiết tính năng',
      details: {
        highlights: [
          'Chọn một cột tiêu chí (ví dụ: Tên chi nhánh, Khu vực, Tháng) để tách thành các file riêng.',
          'Tự động giữ nguyên định dạng tiêu đề, màu sắc, font chữ của bảng gốc.',
          'Tự động đóng gói tất cả các file đã tách thành 1 file .ZIP nén gọn gàng để tải về.',
          'Bảo mật dữ liệu tuyệt đối: Dễ dàng gửi file phân quyền cho từng phòng ban mà không lộ dữ liệu chung.'
        ]
      }
    },
    {
      id: 'diff',
      title: 'So sánh 2 File & Kiểm toán số liệu (Diff)',
      category: 'matching',
      categoryName: 'Đối chiếu & Nhập liệu',
      status: 'upcoming',
      statusText: 'Sắp ra mắt • Q2/2026',
      icon: '⚖️',
      iconTheme: 'icon-blue',
      featured: false,
      desc: 'Đặt 2 phiên bản file bảng tính cạnh nhau (bản cũ vs bản mới), tự động rà soát từng ô số liệu và tô màu highlight các ô bị chỉnh sửa, dòng thêm mới hoặc bị xóa.',
      tags: ['Tô màu ô lệch', 'So sánh phiên bản', 'Kiểm toán số', 'Báo cáo sai khác'],
      actionText: 'Xem chi tiết tính năng',
      details: {
        highlights: [
          'So sánh từng ô giá trị giữa 2 file Excel có cùng cấu trúc hoặc khác nhau số dòng.',
          'Mã màu trực quan: Xanh lá (Dòng mới), Đỏ (Dòng bị xóa), Vàng cam (Ô bị thay đổi số liệu).',
          'Xem trước bảng khác biệt song song (Side-by-side diff view).',
          'Xuất báo cáo sai lệch kèm nhật ký ô cũ -> ô mới cho kiểm toán và sếp duyệt.'
        ]
      }
    },
    {
      id: 'cleaner',
      title: 'Làm sạch & Chuẩn hóa dữ liệu',
      category: 'cleaning',
      categoryName: 'Xử lý & Làm sạch',
      status: 'upcoming',
      statusText: 'Sắp ra mắt • Q2/2026',
      icon: '✨',
      iconTheme: 'icon-emerald',
      featured: false,
      desc: 'Tự động xóa khoảng trắng dư thừa, chuẩn hóa định dạng ngày tháng (DD/MM/YYYY), đồng bộ viết hoa/thường, xóa dòng dữ liệu trùng lặp và chuyển đổi số tiền thành chữ tiếng Việt.',
      tags: ['Chuẩn hóa ngày', 'Xóa trùng lặp', 'Trim khoảng trắng', 'Đọc số ra chữ'],
      actionText: 'Xem chi tiết tính năng',
      details: {
        highlights: [
          'Tự động xóa khoảng trắng ẩn ở đầu/cuối ô và khoảng trắng kép giữa các từ.',
          'Chuyển đổi các định dạng ngày tháng lộn xộn (M/D/Y, YMD, text) về chuẩn Việt Nam (DD/MM/YYYY).',
          'Xóa triệt để các dòng trùng lặp dựa trên 1 hoặc nhiều cột định danh.',
          'Tích hợp thuật toán chuyển đổi số tiền thành chữ tiếng Việt chuẩn quy chuẩn kế toán.'
        ]
      }
    },
    {
      id: 'converter',
      title: 'Chuyển đổi Excel ⮂ CSV, JSON, PDF',
      category: 'utilities',
      categoryName: 'Tiện ích & Sửa lỗi',
      status: 'upcoming',
      statusText: 'Sắp ra mắt • Q2/2026',
      icon: '🔁',
      iconTheme: 'icon-purple',
      featured: false,
      desc: 'Chuyển đổi siêu tốc file Excel sang CSV chuẩn UTF-8 (không lỗi font tiếng Việt khi mở trên hệ thống khác), kết xuất JSON cho lập trình viên hoặc xuất bản in PDF sắc nét.',
      tags: ['Excel sang CSV', 'Xuất JSON', 'Chuẩn UTF-8', 'In ấn PDF'],
      actionText: 'Xem chi tiết tính năng',
      details: {
        highlights: [
          'Xuất CSV chuẩn UTF-8 với BOM, mở trực tiếp trên Excel không bị lỗi font tiếng Việt.',
          'Chuyển đổi bảng tính thành cấu trúc mảng JSON lồng nhau hoặc phẳng phục vụ API.',
          'Xem trước trang in và xuất PDF định dạng trang A4/A3 vừa khít, không bị tràn lề.'
        ]
      }
    },
    {
      id: 'lookup',
      title: 'Tra cứu & Trích xuất siêu tốc (Smart VLOOKUP)',
      category: 'utilities',
      categoryName: 'Tiện ích & Sửa lỗi',
      status: 'upcoming',
      statusText: 'Sắp ra mắt • Q2/2026',
      icon: '🎯',
      iconTheme: 'icon-amber',
      featured: false,
      desc: 'Thay thế các hàm VLOOKUP / XLOOKUP cồng kềnh. Tra cứu dữ liệu kết hợp 2-3 điều kiện giữa các bảng tính lớn hàng trăm nghìn dòng mà không gây treo máy hay crash Excel.',
      tags: ['Đa điều kiện', 'Không lag máy', 'Tra cứu tức thì', 'Trích xuất cột'],
      actionText: 'Xem chi tiết tính năng',
      details: {
        highlights: [
          'Tra cứu kết hợp nhiều điều kiện (ví dụ: vừa trùng Mã nhân viên VÀ vừa trùng Tháng).',
          'Tốc độ cực nhanh nhờ chỉ mục băm (Hash Index) trong bộ nhớ, xử lý 100,000 dòng dưới 1 giây.',
          'Trích xuất tự động các cột tương ứng sang bảng kết quả mà không cần gõ công thức.'
        ]
      }
    },
    {
      id: 'doctor',
      title: 'Bác sĩ Công thức & Kiểm tra lỗi Excel',
      category: 'utilities',
      categoryName: 'Tiện ích & Sửa lỗi',
      status: 'upcoming',
      statusText: 'Sắp ra mắt • Q2/2026',
      icon: '🩺',
      iconTheme: 'icon-rose',
      featured: false,
      desc: 'Tự động quét toàn bộ file bảng tính để phát hiện các ô lỗi #REF!, #VALUE!, #N/A, #DIV/0!, các vòng lặp tham chiếu tròn và đề xuất giải pháp khắc phục triệt để.',
      tags: ['Quét lỗi #REF!', 'Tham chiếu tròn', 'Cảnh báo rủi ro', 'Gợi ý sửa'],
      actionText: 'Xem chi tiết tính năng',
      details: {
        highlights: [
          'Liệt kê danh sách tất cả các ô bị lỗi công thức kèm vị trí chính xác (Sheet, Ô).',
          'Giải thích nguyên nhân lỗi bằng tiếng Việt dễ hiểu và hướng dẫn cách sửa.',
          'Phát hiện các tham chiếu tròn (Circular References) gây chậm hoặc đơ file Excel.'
        ]
      }
    }
  ];

  // DOM Elements
  const grid = document.getElementById('efAppGrid');
  const searchInput = document.getElementById('efHubSearch');
  const filterTabs = document.getElementById('efFilterTabs');

  let currentCategory = 'all';
  let searchQuery = '';

  // Render App Grid
  function renderAppGrid() {
    if (!grid) return;

    const filtered = TOOLS.filter(tool => {
      const matchCat = (currentCategory === 'all') ||
        (currentCategory === 'active' && tool.status === 'active') ||
        (tool.category === currentCategory);

      const q = searchQuery.trim().toLowerCase();
      const matchQuery = !q ||
        tool.title.toLowerCase().includes(q) ||
        tool.desc.toLowerCase().includes(q) ||
        tool.tags.some(t => t.toLowerCase().includes(q)) ||
        tool.categoryName.toLowerCase().includes(q);

      return matchCat && matchQuery;
    });

    if (filtered.length === 0) {
      grid.innerHTML = `
        <div style="grid-column: 1/-1; text-align: center; padding: 48px 20px; background: #fff; border-radius: 16px; border: 1px dashed #cbd5e1;">
          <div style="font-size: 32px; margin-bottom: 10px;">🔍</div>
          <h3 style="margin: 0 0 6px; font-size: 17px; color: #1e293b;">Không tìm thấy công cụ phù hợp</h3>
          <p style="margin: 0; color: #64748b; font-size: 13.5px;">Thử tìm kiếm với từ khóa khác như "ghép", "sao kê", "tách", "so sánh", "làm sạch"...</p>
        </div>
      `;
      return;
    }

    grid.innerHTML = filtered.map(tool => {
      const isFeatured = tool.featured;
      const isUpcoming = tool.status === 'upcoming';

      return `
        <div class="ef-app-card ${isFeatured ? 'featured' : ''}" data-tool-id="${tool.id}">
          <div class="ef-card-header">
            <div class="ef-card-icon-box ${tool.iconTheme}">
              ${tool.icon}
            </div>
            <span class="ef-status-badge ${tool.status}">
              ${isUpcoming ? '🕒 ' : '🟢 '}${tool.statusText}
            </span>
          </div>
          <div class="ef-card-category">${tool.categoryName}</div>
          <h3 class="ef-card-title">${tool.title}</h3>
          <p class="ef-card-desc">${tool.desc}</p>
          <div class="ef-card-tags">
            ${tool.tags.map(t => `<span class="ef-tag">#${t}</span>`).join('')}
          </div>
          <div class="ef-card-action">
            <button type="button" class="ef-launch-btn ${isFeatured ? 'primary' : 'outline'}" data-action-tool="${tool.id}">
              ${tool.actionText}
            </button>
          </div>
        </div>
      `;
    }).join('');

    // Attach click events
    grid.querySelectorAll('[data-action-tool]').forEach(btn => {
      btn.addEventListener('click', () => {
        const toolId = btn.dataset.actionTool;
        handleToolAction(toolId);
      });
    });
  }

  // Handle click on tool action
  function handleToolAction(toolId) {
    if (toolId === 'excelflow') {
      // Chuyển sang màn hình công cụ ExcelFlow
      if (typeof window.efGoPage === 'function') {
        window.efGoPage('import');
      } else {
        const navImport = document.getElementById('navImport');
        if (navImport) navImport.click();
      }
    } else {
      // Mở modal giới thiệu chi tiết công cụ sắp ra mắt
      const tool = TOOLS.find(t => t.id === toolId);
      if (tool) openFeatureModal(tool);
    }
  }

  // Feature Preview Modal
  function openFeatureModal(tool) {
    const modalBackdrop = document.getElementById('efFeatureModalBackdrop');
    const modalTitle = document.getElementById('efFmodalTitle');
    const modalIcon = document.getElementById('efFmodalIcon');
    const modalCategory = document.getElementById('efFmodalCategory');
    const modalDesc = document.getElementById('efFmodalDesc');
    const modalSpecs = document.getElementById('efFmodalSpecs');

    if (!modalBackdrop) return;

    if (modalTitle) modalTitle.textContent = tool.title;
    if (modalIcon) modalIcon.textContent = tool.icon;
    if (modalCategory) modalCategory.textContent = tool.categoryName + ' • ' + tool.statusText;
    if (modalDesc) modalDesc.textContent = tool.desc;

    if (modalSpecs && tool.details && tool.details.highlights) {
      modalSpecs.innerHTML = tool.details.highlights.map(h => `<li>${h}</li>`).join('');
    }

    modalBackdrop.hidden = false;
  }

  function closeFeatureModal() {
    const modalBackdrop = document.getElementById('efFeatureModalBackdrop');
    if (modalBackdrop) modalBackdrop.hidden = true;
  }

  // Setup Event Listeners
  function init() {
    // Search input
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value;
        renderAppGrid();
      });
    }

    // Filter tabs
    if (filterTabs) {
      filterTabs.querySelectorAll('[data-filter]').forEach(btn => {
        btn.addEventListener('click', () => {
          filterTabs.querySelectorAll('[data-filter]').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          currentCategory = btn.dataset.filter;
          renderAppGrid();
        });
      });
    }

    // Modal close events
    const closeBtn = document.getElementById('efFmodalClose');
    const modalBackdrop = document.getElementById('efFeatureModalBackdrop');
    const modalUseActiveBtn = document.getElementById('efFmodalUseActive');

    if (closeBtn) closeBtn.addEventListener('click', closeFeatureModal);
    if (modalBackdrop) {
      modalBackdrop.addEventListener('click', (e) => {
        if (e.target.id === 'efFeatureModalBackdrop') closeFeatureModal();
      });
    }

    if (modalUseActiveBtn) {
      modalUseActiveBtn.addEventListener('click', () => {
        closeFeatureModal();
        handleToolAction('excelflow');
      });
    }

    // ESC to close modal
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeFeatureModal();
    });

    renderAppGrid();
  }

  // Run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
