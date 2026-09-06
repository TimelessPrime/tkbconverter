# 🚀 Hướng Dẫn Nhanh - TKB Converter

**Dành cho người dùng bình thường (không cần biết code)**

---

## 📖 Trong 2 Phút

### Bước 1️⃣: Truy Cập Ứng Dụng
Mở trình duyệt và vào:
```
https://[username].github.io/tkb-converter/
```

Hoặc nếu chạy local:
```
http://localhost:8000
```

### Bước 2️⃣: Chuẩn Bị File Excel
- Có file Excel chứa Thời Khóa Biểu toàn trường
- File phải có cấu trúc:
  - **Cột A**: Thứ (Thứ 2, Thứ 3, ...)
  - **Cột C**: Tiết (Tiết 1, Tiết 2, ...)
  - **Cột D+**: Tên lớp (10A1, 10A2, ...) và tên môn học

### Bước 3️⃣: Tải File
1. Mở ứng dụng
2. Kéo file Excel vào khu vực upload (hoặc click để chọn file)
3. Chờ 2-3 giây để ứng dụng xử lý

### Bước 4️⃣: Chọn Lớp
1. Dropdown sẽ hiển thị danh sách lớp được tìm thấy
2. Chọn 1 lớp (VD: Lớp 10A1)

### Bước 5️⃣: Xem Trước
Bảng TKB sẽ hiển thị bao gồm:
- **Buổi Sáng**: Tiết 1-5
- **Buổi Chiều**: Tiết 1-5

Kiểm tra dữ liệu có đúng không.

### Bước 6️⃣: Xuất File Word
Nhấn nút **"📥 Tải File Word"**

File sẽ được tải về với tên: `TKB_Lop_10A1.docx` (tùy lớp chọn)

### ✅ Hoàn Tất!

---

## ❌ Nếu Gặp Vấn Đề

### ❓ Vấn đề 1: File không tải được

**Nguyên nhân:**
- File không phải Excel (.xlsx) hoặc (.xls)
- File bị hỏng

**Cách khắc phục:**
1. Kiểm tra file là `.xlsx` hoặc `.xls`
2. Mở file bằng Excel để kiểm tra dữ liệu
3. Lưu lại file dưới dạng `.xlsx`

---

### ❓ Vấn đề 2: Không tìm thấy lớp nào

**Nguyên nhân:**
- Tên lớp không đúng format (không phải 10A1, 10A2, ...)
- File Excel cấu trúc sai (Thứ, Tiết, Tên lớp không ở đúng cột)

**Cách khắc phục:**
1. Kiểm tra tên lớp trong file: 10A1, 10A2, 10A3, ... (không phải 10-A-1)
2. Kiểm tra cột:
   - Cột A: Thứ 2, Thứ 3, ..., Thứ 7
   - Cột C: Tiết 1, Tiết 2, ..., Tiết 5
   - Cột D+: Tên lớp và tên môn
3. Tham khảo file mẫu (xem **SAMPLE_DATA.md**)

---

### ❓ Vấn đề 3: Bảng xuất hiện nhưng dữ liệu sai

**Nguyên nhân:**
- Cấu trúc file Excel không đúng
- Dữ liệu sắp xếp sai trong file

**Cách khắc phục:**
1. Kiểm tra file Excel:
   - Mỗi dòng có: Thứ + Tiết + Tên môn không?
   - Tên môn có kèm tên giáo viên không? (VD: "Toán-Minh" là đúng)
2. So sánh với file mẫu trong **SAMPLE_DATA.md**

---

### ❓ Vấn đề 4: File Word được tải về nhưng không mở được

**Nguyên nhân:**
- Trình duyệt không lưu file đúng
- Antivirus chặn file

**Cách khắc phục:**
1. Kiểm tra thư mục "Downloads"
2. Thử download lại
3. Nếu vẫn không được, thử dùng trình duyệt khác

---

## 💡 Mẹo Sử Dụng

### Tip 1: Sao chép bảng từ Word sang Excel
1. Mở file Word được xuất
2. Chọn bảng (Ctrl+A trong bảng)
3. Copy (Ctrl+C)
4. Paste vào Excel (Ctrl+V)

### Tip 2: In bảng TKB
1. Mở file Word được xuất
2. Ctrl+P (Print)
3. Chọn máy in
4. In

### Tip 3: Sửa bảng trong Word
1. Mở file Word được xuất
2. Double-click vào ô để chỉnh sửa
3. Lưu file (Ctrl+S)

### Tip 4: Tạo file Excel mẫu
Nếu bạn chưa có file Excel thích hợp:
1. Mở **SAMPLE_DATA.md** trong repository
2. Copy dữ liệu mẫu
3. Paste vào Excel
4. Sửa tên giáo viên và tên lớp theo thực tế
5. Lưu file

---

## 🆘 Cần Giúp Đỡ Thêm?

1. **Đọc file README.md**: Có hướng dẫn chi tiết
2. **Kiểm tra SAMPLE_DATA.md**: Xem cấu trúc file Excel đúng
3. **Mở Console (F12)**: Xem có error gì không
4. **Liên hệ**: Tạo Issue trên GitHub repository

---

## ✨ Các Bước Để Thành Công

- [ ] Tải ứng dụng (hoặc truy cập link nếu đã deploy)
- [ ] Chuẩn bị file Excel với cấu trúc đúng
- [ ] Upload file
- [ ] Chọn lớp
- [ ] Xem trước bảng
- [ ] Xuất file Word
- [ ] ✅ Hoàn tất!

---

**Chúc bạn sử dụng thành công! 🎉**

Nếu có câu hỏi, tạo Issue trên GitHub hoặc liên hệ với quản trị viên.