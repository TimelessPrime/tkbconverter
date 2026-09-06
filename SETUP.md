# 🔧 Hướng Dẫn Cài Đặt Chi Tiết TKB Converter

## 📋 Yêu Cầu Tiên Quyết

- Git (tùy chọn, nếu muốn dùng version control)
- Trình duyệt hiện đại (Chrome, Firefox, Safari, Edge)
- Tài khoản GitHub (nếu muốn deploy trên GitHub Pages)

## 🚀 Cách 1: Deploy trên GitHub Pages (Khuyên Dùng)

### Bước 1: Tạo Repository GitHub

1. Truy cập https://github.com/new
2. Tạo repository mới với tên: `tkb-converter`
3. Chọn "Public"
4. Click "Create repository"

### Bước 2: Clone Repository

Mở Terminal/Command Prompt và chạy:

```bash
git clone https://github.com/[username]/tkb-converter.git
cd tkb-converter
```

Thay `[username]` bằng username GitHub của bạn.

### Bước 3: Thêm Files

Copy 3 file chính vào thư mục project:
- `index.html`
- `style.css`
- `app.js`

Cấu trúc thư mục sẽ như sau:
```
tkb-converter/
├── index.html
├── style.css
├── app.js
├── README.md
├── SETUP.md
├── .gitignore
└── .github/
    └── workflows/
        └── deploy.yml
```

### Bước 4: Commit & Push

Chạy các lệnh sau:

```bash
git add .
git commit -m "Initial commit: Add TKB Converter app"
git push origin main
```

### Bước 5: Kích Hoạt GitHub Pages

1. Vào repository trên GitHub
2. Chọn **Settings** → **Pages**
3. Chọn:
   - Source: **Deploy from a branch**
   - Branch: **main**
   - Folder: **/(root)**
4. Click **Save**

GitHub sẽ tự động deploy. Sau vài phút (3-5 phút), ứng dụng sẽ có sẵn tại:

```
https://[username].github.io/tkb-converter/
```

**Hoàn tất! 🎉**

---

## 🖥️ Cách 2: Chạy Locally trên Máy Tính

### Cách 2A: Sử dụng Python

Nếu bạn có Python cài đặt:

1. Lưu 3 file chính vào 1 thư mục
2. Mở Terminal/Command Prompt tại thư mục đó
3. Chạy lệnh:

```bash
# Python 3
python -m http.server 8000

# Hoặc Python 2
python -m SimpleHTTPServer 8000
```

4. Mở trình duyệt và truy cập: `http://localhost:8000`

### Cách 2B: Sử dụng Node.js

Nếu bạn có Node.js:

1. Cài đặt `http-server` globally:

```bash
npm install -g http-server
```

2. Chạy:

```bash
http-server
```

3. Truy cập theo địa chỉ hiển thị (thường là `http://127.0.0.1:8080`)

### Cách 2C: Sử dụng Live Server (VS Code)

Nếu bạn dùng VS Code:

1. Cài đặt extension "Live Server"
2. Click chuột phải vào `index.html`
3. Chọn "Open with Live Server"
4. Trình duyệt sẽ mở tự động

---

## 📝 Cấu Trúc File Tạo Ra

Sau khi clone, project của bạn sẽ có:

```
tkb-converter/
│
├── index.html              # Trang HTML chính
│   └── Chứa giao diện UI (Upload, Dropdown, Preview, Export)
│
├── style.css               # Stylesheet
│   └── Định dạng giao diện, responsive, animation
│
├── app.js                  # Logic JavaScript chính
│   ├── Class TKBConverter
│   ├── Xử lý Upload File Excel
│   ├── Tìm kiếm Lớp (Multi-block scanning)
│   ├── Làm sạch Dữ liệu
│   ├── Xây Bảng 2D
│   └── Xuất File Word (.docx)
│
├── README.md               # Tài liệu chính
├── SETUP.md                # Tài liệu này (Hướng dẫn cài đặt)
├── .gitignore              # Git ignore file
└── .github/
    └── workflows/
        └── deploy.yml      # GitHub Actions (tự động deploy)
```

---

## 🔍 Xác Minh Cài Đặt

### Kiểm Tra Xem Ứng Dụng Hoạt Động

1. Tải file Excel mẫu có chứa data TKB
2. Upload file vào ứng dụng
3. Dropdown nên hiển thị danh sách lớp (10A1, 10A2, ..., 10A14)
4. Chọn 1 lớp, xem preview bảng
5. Nhấn nút "Tải File Word" và kiểm tra file được tải về

Nếu tất cả bước trên hoạt động, ✅ **Setup thành công!**

### Nếu Gặp Vấn Đề

#### Issue 1: GitHub Pages không hoạt động
- Kiểm tra Settings → Pages có được kích hoạt không
- Đảm bảo branch đã chọn là `main`
- Chờ 5-10 phút sau khi push
- Xóa cache trình duyệt (Ctrl+Shift+Delete)

#### Issue 2: File Excel không tải được
- Kiểm tra file là `.xlsx` hoặc `.xls`
- Mở Console (F12) xem error message
- File phải có định dạng cột: A (Thứ), C (Tiết), D+ (Lớp)

#### Issue 3: Không tìm thấy lớp nào
- Kiểm tra tên lớp khớp pattern: `10A[số]` (VD: 10A1, 10A2, ..., 10A14)
- Đảm bảo dữ liệu ở đúng vị trí trong file Excel
- Mở Console để xem dữ liệu được trích xuất

---

## 🎨 Tuỳ Chỉnh Ứng Dụng

### Thay Đổi Tên Ứng Dụng

Chỉnh sửa trong `index.html`:
```html
<h1>📅 TKB Converter</h1>
<p class="subtitle">Hệ Thống Lọc & Chuyển Đổi Thời Khóa Biểu Cá Nhân</p>
```

### Thay Đổi Màu Sắc

Chỉnh sửa trong `style.css`:
```css
/* Gradient chính */
background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);

/* Hoặc chỉnh màu riêng biệt */
color: #667eea; /* Màu xanh tím */
```

### Thêm Chuẩn Hóa Môn Học Mới

Chỉnh sửa trong `app.js`, hàm `cleanSubjectName()`:
```javascript
const replacements = {
    'Chào cờ': 'SHDC',
    'HĐTNHN': 'HĐ TNHN',
    'Môn Mới': 'TT', // Thêm dòng này
};
```

### Thay Đổi Ngày Học

Chỉnh sửa trong `app.js`:
```javascript
this.days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
// Hoặc nếu không có Thứ 7:
this.days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];
```

---

## 📊 Cấu Trúc File Excel Mong Đợi

Ứng dụng mong đợi file Excel có cấu trúc như sau:

| Cột A | Cột B | Cột C   | Cột D | Cột E | Cột F |
|-------|-------|---------|-------|-------|-------|
| Thứ   | -     | Tiết    | 10A1  | 10A2  | 10A3  |
| Thứ 2 | -     | Tiết 1  | SHDC  | Toán  | Anh   |
| Thứ 2 | -     | Tiết 2  | -     | Toán  | Anh   |
| Thứ 3 | -     | Tiết 1  | Anh   | Anh   | Toán  |
| ...   | ...   | ...     | ...   | ...   | ...   |

**Quan trọng:**
- Cột A: Tên Thứ (Thứ 2, Thứ 3, ..., Thứ 7)
- Cột C: Tiết học (Tiết 1, Tiết 2, ..., Tiết 5)
- Cột D+: Tên lớp và dữ liệu môn học
- Dữ liệu có thể chứa tên giáo viên (sẽ bị loại bỏ tự động)

---

## 🔒 Bảo Mật

- ✅ **100% Client-side**: Không có server, dữ liệu không gửi đi đâu
- ✅ **File local**: Tất cả xử lý trên máy tính của bạn
- ✅ **Open Source**: Code công khai, bạn có thể tự kiểm tra

---

## 📚 Tài Nguyên Bổ Sung

### Thư Viện Sử Dụng
- **SheetJS**: https://sheetjs.com/
- **docx.js**: https://docx.js.org/
- **FileSaver.js**: https://github.com/eligrey/FileSaver.js

### Hỗ Trợ GitHub Pages
- https://docs.github.com/en/pages

### Hỗ Trợ Git
- https://git-scm.com/doc

---

## ✅ Checklist Deploy

- [ ] Repository tạo trên GitHub
- [ ] Clone repository về máy
- [ ] Thêm 3 file chính (index.html, style.css, app.js)
- [ ] Commit & push lên GitHub
- [ ] Kích hoạt GitHub Pages (Settings → Pages)
- [ ] Chờ 5-10 phút để deploy hoàn tất
- [ ] Truy cập URL GitHub Pages
- [ ] Test với file Excel mẫu
- [ ] ✅ Hoàn tất!

---

## 🎯 Bước Tiếp Theo

Sau khi deploy xong:

1. **Chia sẻ URL**: Chia sẻ link GitHub Pages cho bạn bè/đồng nghiệp
2. **Tạo tài liệu**: Viết hướng dẫn sử dụng cho người dùng
3. **Lấy feedback**: Thu thập ý kiến để cải thiện
4. **Phát triển thêm**: Thêm tính năng mới nếu cần

---

## 📞 Cần Giúp Đỡ?

- Mở **Issue** trên GitHub repository
- Kiểm tra **Console** (F12) để xem error messages
- Đọc **README.md** để hiểu rõ hơn logic

---

**Good luck! 🚀**

Chúc bạn cài đặt thành công TKB Converter!