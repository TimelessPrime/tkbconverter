# 📅 TKB Converter - Chuyển Đổi Thời Khóa Biểu Cá Nhân

Một ứng dụng web đơn giản, chạy **100% Client-side**, hỗ trợ chuyển đổi dữ liệu Thời Khóa Biểu (TKB) thô toàn trường dạng **Ma trận dọc** thành **Bảng TKB Ngang Buổi Chiều** theo từng lớp và xuất ra file **Excel (`.xlsx`)** chuẩn khung định dạng.

---

## 🚀 Đặc Điểm

- ✅ **100% Client-side**: Hoạt động hoàn toàn trên trình duyệt, bảo mật dữ liệu, không cần backend server.
- ✅ **Hỗ trợ Upload & Fetch URL**: Nhận file Excel qua thao tác Kéo-Thả (Drag & Drop), Chọn file trực tiếp hoặc Tải qua đường link URL.
- ✅ **Trích xuất ngày tự động**: Tự động quét và lấy **Ngày áp dụng** từ tên file upload (Ví dụ: `TKB_05_09_2026.xlsx` -> `05/09/2026`).
- ✅ **Nhận diện lớp động**: Tự động phát hiện và lọc danh sách lớp (định dạng `10A...`).
- ✅ **Xem trước trực quan**: Hiển thị bảng TKB Buổi Chiều (Tiết 1 đến Tiết 5, từ Thứ 2 đến Thứ 7) trước khi xuất file.
- ✅ **Xuất Excel chuẩn khung**: Tạo file `.xlsx` được gộp ô (Merged Cells), căn chỉnh tiêu đề chính xác theo mẫu.
- ✅ **Tối ưu di động & Android**: Tích hợp cơ chế Blob Fallback giúp tải file mượt mà trên Chrome, Edge, Samsung Internet di động.
- ✅ **Debug Console**: Tích hợp công cụ xem log/lỗi trực tiếp ngay trên giao diện web.

---

## 🛠️ Công Nghệ Sử Dụng

### Frontend
- **HTML5, CSS3, JavaScript (ES6+)**

### Thư viện (CDN)
- **SheetJS (`xlsx.full.min.js`)**: Đọc file Excel đầu vào và khởi tạo/ghi file Excel đầu ra.

### Hosting / Deployment
- **GitHub Pages / GitHub Actions**: Tự động deploy CI/CD qua branch `main`.

---

## 📋 Yêu Cầu File Input

### Cấu Trúc Excel Mong Đợi
- **Cột A (Cột 0)**: Thứ (`Thứ 2`, `Thứ 3`, ..., `Thứ 7`)
- **Cột C (Cột 2)**: Tiết (`Tiết 1`, `Tiết 2`, ..., `Tiết 5`)
- **Cột D+ (Cột 3 trở đi)**: Tên Lớp (`10A1`, `10A2`, ..., `10A14`)

Ví dụ cấu trúc:

```text
Thứ  | [trống] | Tiết   | 10A1    | 10A2      | 10A3
-----+---------+--------+---------+-----------+-------
Thứ 2| [trống] | Tiết 1 | SHDC    | Tiếng Anh | Toán
Thứ 2| [trống] | Tiết 2 | HĐ TNHN | Tiếng Anh | Toán
...
```

---

## 🎯 Hướng Dẫn Sử Dụng

### 1. Mở Ứng Dụng
Truy cập trang web (hoặc mở trực tiếp file `index.html` trên trình duyệt):

```
https://timelessprime.github.io/tkbconverter/
```

### 2. Tải File Excel
Kéo file Excel vào khu vực Upload, chọn file từ thiết bị, hoặc dán đường link URL file Excel và nhấn **Tải từ URL**.

### 3. Chọn Lớp
Danh sách Dropdown sẽ hiển thị danh sách các lớp `10A` được tự động tìm thấy.

Chọn lớp cần xuất TKB.

### 4. Xem Trước
Khung xem trước sẽ hiển thị bảng Thời Khóa Biểu Buổi Chiều từ Thứ 2 đến Thứ 7 (Tiết 1 - Tiết 5).

### 5. Xuất File Excel
Nhấn nút **"📥 Xuất File Excel TKB"**.

File sẽ tự động tải về với tên dạng: `TKB_Lop_[TÊN_LỚP].xlsx`.

---

## 📁 Cấu Trúc Thư Mục

```
WebLocGioHoc/
├── app.js              # Logic trích xuất & xử lý file Excel
├── index.html          # Trang giao diện chính
├── README.md           # Tài liệu dự án
└── style.css           # Định dạng giao diện & Responsive design
```

---

## 🔧 Cài Đặt & Deploy

### Option 1: Deploy trên GitHub Pages

1. Fork hoặc tạo Repository mới trên GitHub.
2. Clone về máy local:

```bash
git clone https://github.com/timelessprime/tkbconverter.git
cd tkbconverter
```

3. Push mã nguồn lên GitHub:

```bash
git add .
git commit -m "Deploy WebLocGioHoc App"
git push origin main
```

4. Kích hoạt GitHub Pages trong mục **Settings → Pages** (Chọn branch `main`, Folder `/(root)`).

### Option 2: Chạy Local

Mở trực tiếp file `index.html` bằng bất kỳ trình duyệt nào (Chrome, Edge, Firefox, Safari).

Hoặc chạy local server bằng Python:

```bash
python -m http.server 8000
# Truy cập: http://localhost:8000
```

---

## 🔄 Logic Xử Lý Dữ Liệu

**Nhận Diện Ngày Áp Dụng:**
- Sử dụng Regular Expression để tìm định dạng ngày (`DD.MM.YYYY`, `DD_MM_YYYY`, `DD-MM-YYYY`) trong tên file upload.

**Tìm Kiếm Vị Trí Lớp:**
- Quét qua bảng dữ liệu sheet Excel để tìm cột chứa mã lớp khớp pattern `10A[số]`.

**Làm Sạch Dữ Liệu Môn Học:**
- Loại bỏ phần tên giáo viên (sau dấu `-`).
- Chuẩn hóa tên môn:
  - `Chào cờ` → `SHDC`
  - `HĐTNHN` → `HĐ TNHN`
  - `ND GDĐP` → `GDĐP`

**Xây Dựng Khung Excel Xuất File:**
- Định dạng tiêu đề gộp ô (Merge `A1:H1`): `THỜI KHÓA BIỂU LỚP [TÊN_LỚP_CHUẨN] NĂM HỌC 2026-2027 ÁP DỤNG NGÀY [NGÀY]`
- Định dạng gộp ô cột Buổi (`A3:A7`) hiển thị ký hiệu `C` (Buổi Chiều).
- Xuất dữ liệu mảng 2D gồm 8 cột (Buổi, Tiết, Thứ 2 -> Thứ 7).

---

## 🐛 Gỡ Rối (Troubleshooting)

- **File không tải được**: Kiểm tra lại định dạng file (`.xlsx` hoặc `.xls`) và đảm bảo đúng vị trí các cột Thứ/Tiết.
- **Không tìm thấy lớp**: Kiểm tra tên lớp trong file gốc có bắt đầu bằng `10A` hay không.
- **Không tải được file trên Android**: Bật nút **🐛 Enable Debug** ở góc trên màn hình để kiểm tra log lỗi trực tiếp.

---

## 📄 License

MIT License - Tự do sử dụng, sửa đổi và phân phối.

---

Made by **timeless** with ❤️