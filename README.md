# ExcelFlow Suite

> Nền tảng ứng dụng & tiện ích xử lý dữ liệu Excel chuyên sâu, hoạt động hoàn toàn cục bộ trên trình duyệt (client-side), bảo mật tuyệt đối dữ liệu.

## Bức tranh toàn cảnh Hệ sinh thái (Ecosystem)

1. **⚡ Ghép & Đối chiếu dữ liệu (ExcelFlow)** *(Đang hoạt động • Sẵn sàng sử dụng)*:
   - Tự động đối chiếu theo cột khóa hoặc số dòng, ghép dữ liệu nhiều file nguồn vào file đích.
   - Thuật toán tối ưu cho sao kê ngân hàng (nhập ngược, bỏ dòng tổng cộng, tự tìm dòng trống).
   - Xem trước kết quả, lọc theo trạng thái và hỗ trợ Popup CRUD (thêm, sửa, xóa dòng trực tiếp).
   - Lưu trữ lịch sử an toàn trong IndexedDB, hỗ trợ sao lưu & khôi phục dạng file ZIP.

2. **📑 Gộp nhiều file Excel (Multi-Workbook Merger)** *(Lộ trình Q2/2026)*:
   - Gộp hàng loạt file bảng tính hoặc nhiều Sheet thành một bảng tổng hợp duy nhất.

3. **✂️ Tách file thông minh (Smart Splitter)** *(Lộ trình Q2/2026)*:
   - Tự động tách bảng tính lớn thành nhiều file con theo chi nhánh, tháng, nhân viên và đóng gói ZIP.

4. **🔍 So sánh & Phát hiện lệch số (Diff & Reconcile)** *(Lộ trình Q2/2026)*:
   - Đối chiếu 2 phiên bản file bảng tính, tô màu highlight các ô thay đổi, xuất báo cáo chênh lệch.

5. **🧹 Làm sạch & Chuẩn hóa dữ liệu (Data Cleaner)** *(Lộ trình Q2/2026)*:
   - Xóa khoảng trắng ẩn, chuẩn hóa ngày tháng, xóa dòng trùng lặp, chuyển số thành chữ tiếng Việt.

6. **🔄 Chuyển đổi định dạng đa năng (Universal Converter)** *(Lộ trình Q2/2026)*:
   - Chuyển đổi tức thì giữa Excel sang CSV UTF-8 chuẩn, JSON cho lập trình viên và xuất bản in PDF.

7. **📊 Tra cứu đa điều kiện siêu tốc (Smart VLOOKUP)** *(Lộ trình Q2/2026)*:
   - Tra cứu dữ liệu kết hợp nhiều điều kiện giữa các bảng tính lớn mà không gây treo/đơ Excel.

8. **🛡️ Bác sĩ Công thức & Kiểm tra lỗi (Formula Doctor)** *(Lộ trình Q2/2026)*:
   - Tự động quét và phát hiện các ô lỗi `#REF!`, `#N/A`, `#VALUE!`, tham chiếu vòng tròn.

---

## Cấu trúc dự án
```text
EXcelFlow/
├── index.html              # Giao diện chính (Trang chủ Hub & các công cụ)
├── css/
│   ├── style.css          # CSS lõi ứng dụng ExcelFlow
│   └── home.css           # CSS Trang chủ App Workspace Hub, Cards & Navigation
├── js/
│   ├── home.js            # Logic Trang chủ, tìm kiếm công cụ & modal giới thiệu
│   ├── approval.js        # Hộp thoại xác nhận và điều hướng
│   ├── jszip.min.js       # Thư viện xử lý nén/giải nén file Excel (.xlsx)
│   ├── excel-engine.js    # Lõi xử lý Excel, đối chiếu dữ liệu & IndexedDB
│   ├── preview.js         # Xử lý giao diện xem trước & phân trang
│   └── search-crud.js     # Modal tìm kiếm nâng cao & chỉnh sửa trực tiếp
└── README.md
```

## Triển khai (Deployment)
- **Production Cloudflare Pages**: [https://excelflow-6a7.pages.dev](https://excelflow-6a7.pages.dev)
- **GitHub Repository**: [https://github.com/huypv2002/ExcelFlow](https://github.com/huypv2002/ExcelFlow)
