# ExcelFlow

> Ứng dụng web nhập liệu, đối chiếu và ghép dữ liệu Excel tự động, hoạt động hoàn toàn trên trình duyệt (client-side), bảo mật tuyệt đối dữ liệu.

## Tính năng chính
- **Client-Side 100%**: Xử lý file Excel cục bộ trên trình duyệt bằng JSZip và XML DOM parser, không tải file lên server trung gian.
- **Ghép cột thông minh (Column Mapping)**: Tự động phát hiện và gợi ý ghép cặp các cột tương ứng giữa file nguồn và file đích.
- **Bộ lọc & Xem trước dữ liệu (Preview & Filtering)**: Xem trước bảng dữ liệu kết quả, lọc theo trạng thái (tất cả, sẽ nhập, đã ghép, bỏ qua), phân trang.
- **Tìm kiếm & Chỉnh sửa (Popup Search CRUD)**: Hỗ trợ tìm kiếm nhanh, thêm, sửa, xóa dòng trước khi xuất file.
- **Lưu lịch sử & Bản sao lưu (History & Backup)**: Lưu phiên làm việc và lịch sử thao tác vào IndexedDB trên trình duyệt, hỗ trợ tải về bản sao lưu ZIP và khôi phục khi cần.

## Cấu trúc dự án
```text
EXcelFlow/
├── index.html              # Trang giao diện chính
├── css/
│   └── style.css          # Toàn bộ định kiểu giao diện
├── js/
│   ├── approval.js        # Hộp thoại xác nhận và điều hướng
│   ├── jszip.min.js       # Thư viện xử lý nén/giải nén file Excel (.xlsx)
│   ├── excel-engine.js    # Lõi xử lý Excel, đối chiếu dữ liệu & IndexedDB
│   ├── preview.js         # Xử lý giao diện xem trước & phân trang
│   └── search-crud.js     # Modal tìm kiếm nâng cao & chỉnh sửa trực tiếp
└── README.md
```

## Triển khai (Deployment)
Dự án là ứng dụng Single Page tĩnh, có thể triển khai dễ dàng lên:
- **Cloudflare Pages**: `npx wrangler pages deploy . --project-name excelflow`
- **GitHub Pages**
