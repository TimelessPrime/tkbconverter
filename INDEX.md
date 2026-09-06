# 📑 TKB Converter - Danh Sách File & Hướng Dẫn Bắt Đầu

**Chào mừng bạn đến với TKB Converter! 👋**

Dưới đây là danh sách tất cả các file và cách sử dụng chúng.

---

## 🎯 Bắt Đầu Nhanh (30 giây)

### Bước 1: Chọn Cách Deploy
```
Cách 1: GitHub Pages (Khuyên dùng)
  → Vào SETUP.md → Follow các bước

Cách 2: Chạy Locally (Không cần deploy)
  → Mở index.html bằng trình duyệt
```

### Bước 2: Chuẩn Bị File Excel
```
  → Xem SAMPLE_DATA.md để tạo file mẫu
  → Hoặc dùng file Excel hiện tại (phải có cấu trúc đúng)
```

### Bước 3: Sử Dụng
```
  → Mở ứng dụng
  → Upload file Excel
  → Chọn lớp
  → Xuất Word
  → ✅ Done!
```

**Xem chi tiết:** QUICKSTART.md (cho người dùng bình thường)

---

## 📁 Danh Sách File & Cách Dùng

### 🔴 CORE FILES (Chính - cần thiết)

#### 1. **index.html** (3.2 KB)
- **Mục đích**: Trang giao diện chính
- **Nội dung**: 
  - Khu vực upload file
  - Dropdown chọn lớp
  - Khu vực xem trước bảng
  - Nút xuất Word
- **Sử dụng**: Tự động load khi bạn vào ứng dụng
- **Cần sửa?**: Chỉ nếu muốn thay đổi bố cục UI

#### 2. **style.css** (6.7 KB)
- **Mục đích**: Định dạng giao diện
- **Nội dung**:
  - Màu sắc (Gradient purple)
  - Font chữ
  - Layout responsive
  - Animation & transition
- **Sử dụng**: Tự động load từ index.html
- **Cần sửa?**: Nếu muốn thay đổi màu sắc, font chữ, hoặc layout

#### 3. **app.js** (16 KB)
- **Mục đích**: Logic chính của ứng dụng
- **Nội dung**:
  - Upload Excel & parse SheetJS
  - Tìm kiếm lớp (multi-block scanning)
  - Trích xuất & làm sạch dữ liệu
  - Xây dựng bảng 2D
  - Xuất file Word với docx.js
- **Sử dụng**: Tự động load từ index.html
- **Cần sửa?**: Nếu muốn tuỳ chỉnh logic hoặc thêm tính năng

---

### 📘 DOCUMENTATION FILES (Tài liệu)

#### 4. **README.md** (6.1 KB) ⭐ READ ME FIRST
- **Dành cho**: Tất cả mọi người
- **Nội dung chính**:
  - Tổng quan dự án
  - Công nghệ sử dụng
  - Hướng dẫn sử dụng (6 bước)
  - Cách cài đặt & deploy
  - Gỡ rối (troubleshooting)
  - Tuỳ chỉnh (customization)
- **Khi nào đọc**: 
  - Lần đầu tiên bạn dùng dự án
  - Muốn hiểu rõ hơn ứng dụng
  - Gặp vấn đề cần giải quyết
- **Thời gian đọc**: 10-15 phút

#### 5. **QUICKSTART.md** (4.3 KB) ⭐ FOR END USERS
- **Dành cho**: Người dùng bình thường (không cần code)
- **Nội dung chính**:
  - Hướng dẫn 6 bước nhanh
  - FAQ & troubleshooting đơn giản
  - Tips & tricks
- **Khi nào đọc**: 
  - Bạn chỉ muốn biết cách dùng
  - Chưa quen công nghệ
  - Gặp vấn đề cơ bản
- **Thời gian đọc**: 5 phút

#### 6. **SETUP.md** (7.9 KB) ⭐ FOR DEVELOPERS
- **Dành cho**: Developer, người muốn deploy
- **Nội dung chính**:
  - Cài đặt chi tiết từng bước
  - GitHub Pages deployment
  - Local development
  - Tuỳ chỉnh ứng dụng
  - Troubleshooting nâng cao
- **Khi nào đọc**: 
  - Bạn muốn deploy lên GitHub Pages
  - Muốn chạy ứng dụng locally
  - Muốn hiểu kỹ hơn cấu trúc
- **Thời gian đọc**: 15-20 phút

#### 7. **SAMPLE_DATA.md** (9.5 KB) ⭐ FOR DATA SETUP
- **Dành cho**: Người quản lý dữ liệu, người tạo file Excel
- **Nội dung chính**:
  - Cấu trúc file Excel mong đợi
  - Ví dụ dữ liệu cụ thể
  - Cách tạo file Excel mẫu
  - Copy-paste ready data
  - Quy ước tên môn học
  - Cách test
- **Khi nào đọc**: 
  - Bạn muốn tạo file Excel mẫu
  - Không biết cấu trúc file nên sao
  - Muốn test ứng dụng
- **Thời gian đọc**: 10 phút

#### 8. **PROJECT_SUMMARY.md** (Tài liệu này)
- **Dành cho**: Quản lý dự án, team lead
- **Nội dung chính**:
  - Tóm tắt toàn dự án
  - Kiến trúc kỹ thuật
  - Luồng xử lý dữ liệu
  - Technology stack
  - Performance metrics
  - Roadmap phát triển
- **Khi nào đọc**: 
  - Muốn hiểu sâu về dự án
  - Là quản lý/lead developer
  - Muốn plan phát triển tiếp
- **Thời gian đọc**: 20 phút

---

### ⚙️ CONFIGURATION FILES

#### 9. **.gitignore**
- **Mục đích**: Loại trừ file không cần từ Git
- **Nội dung**: 
  - OS files (.DS_Store, Thumbs.db)
  - IDE files (.vscode, .idea)
  - Node modules, build files
- **Sử dụng**: Tự động áp dụng khi push lên GitHub

#### 10. **.github/workflows/deploy.yml**
- **Mục đích**: GitHub Actions - tự động deploy GitHub Pages
- **Nội dung**: CI/CD workflow
- **Sử dụng**: Tự động chạy khi push lên main branch
- **Cần sửa?**: Không (setup tự động)

#### 11. **INDEX.md** (File này)
- **Mục đích**: Danh sách file & hướng dẫn
- **Sử dụng**: Bạn đang đọc đây 😄

---

## 🎓 Lộ Trình Đọc Theo Vai Trò

### 👤 Nếu Bạn Là: **Người Dùng Bình Thường**
```
1. Đọc QUICKSTART.md (5 phút)
   ↓
2. Mở ứng dụng & xử lý
   ↓
3. Gặp vấn đề? → Đọc QUICKSTART.md FAQ
   ↓
4. Vẫn không giải quyết? → Đọc README.md Troubleshooting
```

### 💻 Nếu Bạn Là: **Developer / Người Deploy**
```
1. Đọc README.md (10 phút)
   ↓
2. Đọc SETUP.md (15 phút)
   ↓
3. Follow SETUP.md Deploy Steps
   ↓
4. Test & Verify
   ↓
5. Tuỳ chỉnh nếu cần (sửa app.js, style.css)
```

### 📊 Nếu Bạn Là: **Data Manager / Excel Specialist**
```
1. Đọc SAMPLE_DATA.md (10 phút)
   ↓
2. Tạo / Chuẩn bị file Excel theo cấu trúc
   ↓
3. Test với ứng dụng
   ↓
4. Tạo file mẫu cho team khác
```

### 👔 Nếu Bạn Là: **Project Manager / Team Lead**
```
1. Đọc README.md (10 phút)
   ↓
2. Đọc PROJECT_SUMMARY.md (20 phút)
   ↓
3. Hiểu toàn bộ dự án
   ↓
4. Plan phát triển & maintain
   ↓
5. Tham khảo Roadmap trong PROJECT_SUMMARY.md
```

---

## ❓ Bạn Đang Gặp Vấn Đề?

### "Không biết bắt đầu từ đâu"
→ Đọc **QUICKSTART.md** (5 phút)

### "Muốn deploy trên GitHub Pages"
→ Đọc **SETUP.md** → Phần "GitHub Pages Deployment"

### "Muốn chạy locally"
→ Đọc **SETUP.md** → Phần "Local Development"

### "Không biết tạo file Excel"
→ Đọc **SAMPLE_DATA.md** → Bỏ qua các phần code

### "Muốn thay đổi màu sắc/font"
→ Sửa **style.css** (hoặc SETUP.md → Customization)

### "Muốn thêm chức năng mới"
→ Sửa **app.js** (hoặc tham khảo PROJECT_SUMMARY.md → Architecture)

### "Gặp lỗi"
→ Mở **Console (F12)** xem error → Tham khảo **README.md Troubleshooting**

### "Không tìm thấy lớp"
→ Xem **SAMPLE_DATA.md** để kiểm tra cấu trúc Excel

### "File Word không mở được"
→ Xem **QUICKSTART.md → Issue 4**

---

## 📊 File Tree (Cấu Trúc Thư Mục)

```
tkb-converter/
│
├── 📄 Core Files (Chạy ứng dụng)
│   ├── index.html          (Giao diện)
│   ├── style.css           (Styling)
│   └── app.js              (Logic)
│
├── 📚 Documentation (Tài liệu)
│   ├── README.md           ⭐ Đọc trước
│   ├── QUICKSTART.md       ⭐ Cho người dùng
│   ├── SETUP.md            ⭐ Cho developer
│   ├── SAMPLE_DATA.md      ⭐ Cho data setup
│   ├── PROJECT_SUMMARY.md  (Tóm tắt kỹ thuật)
│   └── INDEX.md            (File này)
│
├── ⚙️ Configuration
│   ├── .gitignore
│   └── .github/
│       └── workflows/
│           └── deploy.yml
│
└── 📁 (Optional) GitHub Specific
    ├── .git/               (Nếu dùng Git)
    └── .github/            (GitHub Actions)
```

---

## 🔗 Liên Kết Nhanh

| Nhu Cầu | Đọc File | Mục Cụ Thể |
|--------|----------|-----------|
| Bắt đầu ngay | QUICKSTART.md | Toàn bộ |
| Hiểu ứng dụng | README.md | "Mục Tiêu Dự Án" |
| Deploy GitHub | SETUP.md | "GitHub Pages" |
| Chạy locally | SETUP.md | "Local Development" |
| Tạo file Excel | SAMPLE_DATA.md | "Cách Tạo File Excel" |
| Gỡ rối | README.md | "Gỡ Rối" |
| Tuỳ chỉnh | SETUP.md | "Tuỳ Chỉnh" |
| Hiểu kiến trúc | PROJECT_SUMMARY.md | "Kiến Trúc" |

---

## ✅ Checklist: "Tôi đã sẵn sàng"

- [ ] Đã đọc README.md
- [ ] Biết ứng dụng làm gì
- [ ] Có file Excel với cấu trúc đúng (hoặc tạo file mẫu)
- [ ] Đã test ứng dụng (hoặc sẵn sàng test)
- [ ] Nếu muốn deploy: Đã đọc SETUP.md
- [ ] Nếu deploy GitHub Pages: Đã follow tất cả steps
- [ ] ✅ Sẵn sàng sử dụng!

---

## 📞 Cần Hỗ Trợ?

**Trước khi liên hệ, hãy:**
1. Đọc file tài liệu liên quan
2. Kiểm tra FAQ/Troubleshooting
3. Mở Console (F12) xem error message
4. So sánh với SAMPLE_DATA.md

**Nếu vẫn không giải quyết được:**
- Tạo Issue trên GitHub repository
- Cung cấp error message & context
- Mô tả vấn đề chi tiết

---

## 🎉 Hoàn Tất Setup?

Nếu bạn đã:
- ✅ Đọc tài liệu phù hợp
- ✅ Setup/Deploy ứng dụng
- ✅ Test với file Excel
- ✅ Xuất được file Word

**Chúc mừng! 🎊 Bạn đã sẵn sàng sử dụng TKB Converter!**

---

## 📝 Notes

- **Tất cả file tài liệu là Markdown** (.md) → Có thể đọc với bất kỳ editor nào
- **Các file code** (HTML, CSS, JS) → Cần editor code hoặc VS Code
- **File Excel mẫu** → Xem SAMPLE_DATA.md để copy-paste data

---

## 🚀 Bước Tiếp Theo

1. **Chọn tài liệu** phù hợp với vai trò của bạn
2. **Đọc** trong 5-20 phút
3. **Thực hành** (deploy hoặc sử dụng)
4. **Chia sẻ** với đội ngũ
5. **Feedback** & suggest cải thiện

---

**Happy using TKB Converter! 🎉**

Made with ❤️ by TKB Converter Team

Last Updated: 6 tháng 9, 2026