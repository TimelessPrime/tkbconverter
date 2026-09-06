# 📅 TKB Converter - Chuyển Đổi Thời Khóa Biểu Cá Nhân

Một ứng dụng web đơn giản, chạy **100% Client-side**, hỗ trợ chuyển đổi dữ liệu Thời Khóa Biểu (TKB) thô toàn trường dạng **Ma trận dọc** thành **Bảng TKB Ngang** theo từng lớp và xuất ra file Word (`.docx`).

## 🚀 Đặc Điểm

- ✅ **100% Client-side**: Không cần backend server
- ✅ **Tải file Excel**: Hỗ trợ định dạng `.xlsx` và `.xls`
- ✅ **Chọn lớp động**: Tự động phát hiện tất cả lớp từ file Excel
- ✅ **Xem trước**: Hiển thị bảng TKB trước khi xuất
- ✅ **Xuất Word**: Tạo file `.docx` đẹp và chuyên nghiệp
- ✅ **Responsive Design**: Hoạt động trên desktop, tablet, mobile
- ✅ **GitHub Pages Ready**: Deploy trực tiếp lên GitHub Pages

## 🛠️ Công Nghệ Sử Dụng

### Frontend
- **HTML5, CSS3, JavaScript (ES6+)**

### Thư viện (CDN)
- **SheetJS (xlsx)**: Đọc file Excel
- **docx.js**: Tạo file Word (.docx)
- **FileSaver.js**: Tải file về máy client

### Hosting
- **GitHub Pages**: Miễn phí, không cần server

## 📋 Yêu Cầu File Input

### Format Excel Mong Đợi
- **Cột A**: Thứ (Thứ 2, Thứ 3, ..., Thứ 7)
- **Cột C**: Tiết (Tiết 1, Tiết 2, ..., Tiết 5)
- **Cột D+**: Dữ liệu lớp (10A1, 10A2, ..., 10A14)

Ví dụ cấu trúc:
```
Thứ  | [empty] | Tiết      | 10A1    | 10A2    | 10A3
-----+---------+-----------+---------+---------+-------
Thứ 2| [empty] | Tiết 1    | SHDC    | Tiếng Anh | Toán
Thứ 2| [empty] | Tiết 2    | [empty] | Tiếng Anh | Toán
...
```

## 🎯 Hướng Dẫn Sử Dụng

### 1. Mở Ứng Dụng
Truy cập trang web (hoặc `index.html` nếu chạy locally):
```
https://[username].github.io/tkb-converter/
```

### 2. Tải File Excel
- Kéo file Excel vào khu vực upload, hoặc
- Nhấn vào khu vực để chọn file từ máy

### 3. Chọn Lớp
- Dropdown sẽ hiển thị danh sách tất cả lớp được phát hiện
- Chọn 1 lớp từ danh sách

### 4. Xem Trước
- Bảng TKB sẽ hiển thị bao gồm Buổi Sáng và Buổi Chiều
- Kiểm tra dữ liệu trước khi xuất

### 5. Xuất File Word
- Nhấn nút "📥 Tải File Word"
- File sẽ được tải về với tên: `TKB_Lop_[TÊN_LỚP].docx`

## 📁 Cấu Trúc Thư Mục

```
tkb-converter/
├── index.html          # Trang chính
├── style.css           # Định dạng giao diện
├── app.js              # Logic chính
├── README.md           # Tài liệu này
└── .github/
    └── workflows/
        └── deploy.yml  # (Optional) CI/CD cho GitHub Pages
```

## 🔧 Cài Đặt & Deploy

### Option 1: Deploy trên GitHub Pages
1. Fork repository này hoặc tạo repository mới
2. Clone về máy:
   ```bash
   git clone https://github.com/[username]/tkb-converter.git
   cd tkb-converter
   ```
3. Copy 3 file chính vào thư mục:
   - `index.html`
   - `style.css`
   - `app.js`
4. Push lên GitHub:
   ```bash
   git add .
   git commit -m "Add TKB Converter app"
   git push origin main
   ```
5. Kích hoạt GitHub Pages:
   - Vào Settings → Pages
   - Chọn Branch: `main`, Folder: `/(root)`
   - Lưu

### Option 2: Chạy Locally
1. Clone repository
2. Mở file `index.html` trong trình duyệt
   - Hoặc sử dụng local server:
   ```bash
   python -m http.server 8000
   # Truy cập: http://localhost:8000
   ```

## 🔄 Logic Xử Lý

### 1. Tìm Kiếm Lớp
- Quét toàn bộ sheet Excel tìm các ô khớp pattern: `10A[số]`
- Lưu danh sách lớp được phát hiện

### 2. Trích Xuất Dữ Liệu
- Định vị cột chứa tên lớp được chọn
- Trích xuất từng dòng: Thứ + Tiết + Tên Môn Học

### 3. Làm Sạch Dữ Liệu
- Loại bỏ tên giáo viên (phần sau dấu `-`)
- Chuẩn hóa tên môn:
  - `Chào cờ` → `SHDC`
  - `HĐTNHN` → `HĐ TNHN`
  - `ND GDĐP` → `ND GDĐP` (giữ nguyên)
  - v.v.

### 4. Xây Dựng Bảng 2D
- Ma trận: **Thứ (2-7) × Tiết (1-5)** cho cả Sáng/Chiều
- Điền dữ liệu đã lọc vào đúng vị trí

### 5. Xuất Word
- Tạo Document với bảng kẻ khung chuẩn
- Ghi tiêu đề: `THỜI KHÓA BIỂU LỚP [TÊN_LỚP]`
- Xuất file `.docx` tự động tải về

## ⚙️ Tuỳ Chỉnh

### Thay Đổi Danh Sách Ngày Học
Trong `app.js`, dòng `this.days`:
```javascript
this.days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
```

### Thêm Chuẩn Hóa Môn Học
Trong hàm `cleanSubjectName()`:
```javascript
const replacements = {
    'Chào cờ': 'SHDC',
    'Môn Mới': 'TT', // Thêm dòng này
    // ...
};
```

### Thay Đổi Màu Sắc
Trong `style.css`:
```css
background: linear-gradient(135deg, #YOUR_COLOR1 0%, #YOUR_COLOR2 100%);
```

## 🐛 Gỡ Rối

### File không tải được
- Đảm bảo file là `.xlsx` hoặc `.xls`
- Kiểm tra định dạng cột (A: Thứ, C: Tiết, D+: Lớp)
- Mở console (F12) để xem chi tiết lỗi

### Không tìm thấy lớp
- Kiểm tra tên lớp có khớp pattern `10A[số]` không
- Mở file Excel và xác nhận tên lớp chính xác

### Bảng xuất sai dữ liệu
- Kiểm tra lại cấu trúc file Excel
- Đảm bảo Thứ, Tiết, và Tên Môn ở đúng cột
- Xem console để xem dữ liệu được trích xuất

## 📝 Lịch Sử Thay Đổi

### v1.0 (Initial Release)
- ✅ Upload file Excel
- ✅ Phát hiện tự động danh sách lớp
- ✅ Xem trước bảng TKB
- ✅ Xuất file Word
- ✅ Làm sạch & chuẩn hóa dữ liệu
- ✅ Responsive design

## 📄 License

MIT License - Tự do sử dụng, sửa đổi và phân phối

## 🤝 Đóng Góp

Nếu có ý tưởng cải thiện:
1. Fork repository
2. Tạo branch feature: `git checkout -b feature/your-feature`
3. Commit thay đổi: `git commit -m 'Add your feature'`
4. Push lên branch: `git push origin feature/your-feature`
5. Tạo Pull Request

## 📞 Liên Hệ

Mở Issue nếu gặp bất kỳ vấn đề nào!

---

**Made with ❤️ by TKB Converter Team**