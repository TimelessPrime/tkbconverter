# 📊 Hướng Dẫn Tạo File Excel Mẫu

Để test ứng dụng TKB Converter, bạn cần tạo file Excel với cấu trúc dữ liệu thích hợp.

## 🎯 Cấu Trúc File Excel

### Định Dạng Cột

| Cột | Tên | Mô Tả |
|-----|-----|-------|
| **A** | Thứ | Ngày trong tuần (Thứ 2, Thứ 3, ..., Thứ 7) |
| **B** | (Trống) | Cột này không sử dụng |
| **C** | Tiết | Tiết học (Tiết 1, Tiết 2, ..., Tiết 5) |
| **D+** | Tên Lớp + Dữ liệu | Tên lớp (10A1, 10A2, etc.) và tên môn học |

---

## 📋 Mẫu Dữ Liệu Cụ Thể

Dưới đây là ví dụ data Excel cho lớp **10A1**, **10A2**, **10A3**:

### Buổi Sáng - Dòng 1-5

```
Thứ   | -     | Tiết      | 10A1                 | 10A2           | 10A3
------|-------|-----------|----------------------|----------------|----------
Thứ 2 | -     | Tiết 1    | Chào cờ-Linh         | Tiếng Anh-Vy   | Toán-Minh
Thứ 2 | -     | Tiết 2    | -                    | Tiếng Anh-Vy   | Toán-Minh
Thứ 2 | -     | Tiết 3    | -                    | -              | -
Thứ 2 | -     | Tiết 4    | -                    | -              | -
Thứ 2 | -     | Tiết 5    | -                    | -              | -
```

### Tiếp Tục - Các Ngày Khác

```
Thứ   | -     | Tiết      | 10A1                 | 10A2           | 10A3
------|-------|-----------|----------------------|----------------|----------
Thứ 3 | -     | Tiết 1    | Tiếng Anh-Vy         | Tiếng Anh-Vy   | Toán-Minh
Thứ 3 | -     | Tiết 2    | Tiếng Anh-Vy         | Tiếng Anh-Vy   | Toán-Minh
Thứ 3 | -     | Tiết 3    | -                    | Tin Học-Hùng   | Tin Học-Hùng
Thứ 3 | -     | Tiết 4    | Tin Học-Hùng        | Tin Học-Hùng   | Vật Lí-Anh
Thứ 3 | -     | Tiết 5    | -                    | -              | -

Thứ 4 | -     | Tiết 1    | HĐTNHN               | Toán-Minh      | Hóa Học-Thanh
Thứ 4 | -     | Tiết 2    | HĐTNHN               | Toán-Minh      | Hóa Học-Thanh
Thứ 4 | -     | Tiết 3    | -                    | Lịch Sử-Mạnh   | Lịch Sử-Mạnh
Thứ 4 | -     | Tiết 4    | Ngữ Văn-Hương        | Ngữ Văn-Hương  | Ngữ Văn-Hương
Thứ 4 | -     | Tiết 5    | -                    | -              | -

Thứ 5 | -     | Tiết 1    | GDTC                 | Hóa Học-Thanh  | Vật Lí-Anh
Thứ 5 | -     | Tiết 2    | GDTC                 | Hóa Học-Thanh  | Vật Lí-Anh
Thứ 5 | -     | Tiết 3    | -                    | -              | -
Thứ 5 | -     | Tiết 4    | -                    | Ngữ Văn-Hương  | Ngữ Văn-Hương
Thứ 5 | -     | Tiết 5    | -                    | -              | -

Thứ 6 | -     | Tiết 1    | -                    | Toán-Minh      | Toán-Minh
Thứ 6 | -     | Tiết 2    | -                    | Sinh Học-Liên  | Sinh Học-Liên
Thứ 6 | -     | Tiết 3    | -                    | Vật Lí-Anh     | Vật Lí-Anh
Thứ 6 | -     | Tiết 4    | -                    | Vật Lí-Anh     | Vật Lí-Anh
Thứ 6 | -     | Tiết 5    | -                    | -              | -

Thứ 7 | -     | Tiết 1    | HĐTNHN               | -              | -
Thứ 7 | -     | Tiết 2    | HĐTNHN               | -              | -
Thứ 7 | -     | Tiết 3    | HĐTNHN               | -              | -
Thứ 7 | -     | Tiết 4    | -                    | -              | -
Thứ 7 | -     | Tiết 5    | -                    | -              | -
```

### Buổi Chiều - Bắt Đầu từ Dòng Tiếp Theo

```
Thứ   | -     | Tiết      | 10A1                 | 10A2           | 10A3
------|-------|-----------|----------------------|----------------|----------
Thứ 2 | -     | C Tiết 1  | Tiếng Anh-Vy         | ND GDĐP-Sơn    | Tiếng Anh-Vy
Thứ 2 | -     | C Tiết 2  | ND GDĐP-Sơn         | Tiếng Anh-Vy   | Tiếng Anh-Vy
Thứ 2 | -     | C Tiết 3  | Sinh Học-Liên       | Sinh Học-Liên  | Sinh Học-Liên
Thứ 2 | -     | C Tiết 4  | Hóa Học-Thanh       | Hóa Học-Thanh  | Hóa Học-Thanh
Thứ 2 | -     | C Tiết 5  | Tin Học-Hùng        | Tin Học-Hùng   | Tin Học-Hùng

... (Tiếp tục cho các ngày khác)
```

---

## 🛠️ Cách Tạo File Excel

### Cách 1: Sử Dụng MS Excel

1. Mở Microsoft Excel
2. Tạo file mới
3. Copy-paste dữ liệu trên vào
4. Lưu file với tên: `TKB_Sample.xlsx`

### Cách 2: Sử Dụng Google Sheets

1. Truy cập https://sheets.google.com
2. Tạo Spreadsheet mới
3. Nhập dữ liệu
4. File → Download → Microsoft Excel (.xlsx)

### Cách 3: Sử Dụng LibreOffice Calc

1. Mở LibreOffice Calc
2. Nhập dữ liệu
3. File → Save As → Format: Microsoft Excel (.xlsx)

---

## ✅ Kiểm Tra Dữ Liệu

Trước khi test ứng dụng, hãy kiểm tra:

- [ ] **Cột A**: Chứa "Thứ 2", "Thứ 3", ..., "Thứ 7" đúng thứ tự
- [ ] **Cột C**: Chứa "Tiết 1", "Tiết 2", ..., "Tiết 5" và "C Tiết 1", ... cho buổi chiều
- [ ] **Cột D, E, F, ...**: Tên lớp (10A1, 10A2, 10A3) và tên môn học
- [ ] **Tên môn học**: Theo format "Tên Môn-Tên Giáo Viên" (VD: "Toán-Minh")
- [ ] **Ô trống**: Dùng "-" thay vì để trống hoàn toàn (tùy chọn)

---

## 🔄 Quy Ước Tên Môn Học

Ứng dụng sẽ **tự động làm sạch** tên môn học:

| Input | Output |
|-------|--------|
| `Chào cờ-Linh` | `SHDC` |
| `Chào cờ` | `SHDC` |
| `HĐTNHN` | `HĐ TNHN` |
| `HĐ TNHN` | `HĐ TNHN` |
| `ND GDĐP-Sơn` | `ND GDĐP` |
| `GDTC` | `GDTC` |
| `SHL` | `SHL` |
| `Toán-Minh` | `Toán` |
| `Tiếng Anh-Vy` | `Tiếng Anh` |
| `Vật Lí-Anh` | `Vật Lí` |

---

## 📊 Dữ Liệu Hoàn Chỉnh (Copy-Paste Ready)

Bạn có thể copy-paste trực tiếp vào Excel:

```
Thứ	_	Tiết	10A1	10A2	10A3
Thứ 2	_	Tiết 1	Chào cờ-Linh	Tiếng Anh-Vy	Toán-Minh
Thứ 2	_	Tiết 2	-	Tiếng Anh-Vy	Toán-Minh
Thứ 2	_	Tiết 3	-	-	-
Thứ 2	_	Tiết 4	-	-	-
Thứ 2	_	Tiết 5	-	-	-
Thứ 3	_	Tiết 1	Tiếng Anh-Vy	Tiếng Anh-Vy	Toán-Minh
Thứ 3	_	Tiết 2	Tiếng Anh-Vy	Tiếng Anh-Vy	Toán-Minh
Thứ 3	_	Tiết 3	-	Tin Học-Hùng	Tin Học-Hùng
Thứ 3	_	Tiết 4	Tin Học-Hùng	Tin Học-Hùng	Vật Lí-Anh
Thứ 3	_	Tiết 5	-	-	-
Thứ 4	_	Tiết 1	HĐTNHN	Toán-Minh	Hóa Học-Thanh
Thứ 4	_	Tiết 2	HĐTNHN	Toán-Minh	Hóa Học-Thanh
Thứ 4	_	Tiết 3	-	Lịch Sử-Mạnh	Lịch Sử-Mạnh
Thứ 4	_	Tiết 4	Ngữ Văn-Hương	Ngữ Văn-Hương	Ngữ Văn-Hương
Thứ 4	_	Tiết 5	-	-	-
Thứ 5	_	Tiết 1	GDTC	Hóa Học-Thanh	Vật Lí-Anh
Thứ 5	_	Tiết 2	GDTC	Hóa Học-Thanh	Vật Lí-Anh
Thứ 5	_	Tiết 3	-	-	-
Thứ 5	_	Tiết 4	-	Ngữ Văn-Hương	Ngữ Văn-Hương
Thứ 5	_	Tiết 5	-	-	-
Thứ 6	_	Tiết 1	-	Toán-Minh	Toán-Minh
Thứ 6	_	Tiết 2	-	Sinh Học-Liên	Sinh Học-Liên
Thứ 6	_	Tiết 3	-	Vật Lí-Anh	Vật Lí-Anh
Thứ 6	_	Tiết 4	-	Vật Lí-Anh	Vật Lí-Anh
Thứ 6	_	Tiết 5	-	-	-
Thứ 7	_	Tiết 1	HĐTNHN	-	-
Thứ 7	_	Tiết 2	HĐTNHN	-	-
Thứ 7	_	Tiết 3	HĐTNHN	-	-
Thứ 7	_	Tiết 4	-	-	-
Thứ 7	_	Tiết 5	-	-	-
Thứ 2	_	C Tiết 1	Tiếng Anh-Vy	ND GDĐP-Sơn	Tiếng Anh-Vy
Thứ 2	_	C Tiết 2	ND GDĐP-Sơn	Tiếng Anh-Vy	Tiếng Anh-Vy
Thứ 2	_	C Tiết 3	Sinh Học-Liên	Sinh Học-Liên	Sinh Học-Liên
Thứ 2	_	C Tiết 4	Hóa Học-Thanh	Hóa Học-Thanh	Hóa Học-Thanh
Thứ 2	_	C Tiết 5	Tin Học-Hùng	Tin Học-Hùng	Tin Học-Hùng
Thứ 3	_	C Tiết 1	Tiếng Anh-Vy	Tiếng Anh-Vy	Toán-Minh
Thứ 3	_	C Tiết 2	Tiếng Anh-Vy	Tiếng Anh-Vy	Toán-Minh
Thứ 3	_	C Tiết 3	Tin Học-Hùng	Sinh Học-Liên	Vật Lí-Anh
Thứ 3	_	C Tiết 4	Tin Học-Hùng	Sinh Học-Liên	Vật Lí-Anh
Thứ 3	_	C Tiết 5	-	-	-
Thứ 4	_	C Tiết 1	Toán-Minh	Toán-Minh	Toán-Minh
Thứ 4	_	C Tiết 2	Toán-Minh	Toán-Minh	Toán-Minh
Thứ 4	_	C Tiết 3	Lịch Sử-Mạnh	Lịch Sử-Mạnh	Lịch Sử-Mạnh
Thứ 4	_	C Tiết 4	Ngữ Văn-Hương	Ngữ Văn-Hương	Ngữ Văn-Hương
Thứ 4	_	C Tiết 5	-	-	-
Thứ 5	_	C Tiết 1	Hóa Học-Thanh	Hóa Học-Thanh	Hóa Học-Thanh
Thứ 5	_	C Tiết 2	Hóa Học-Thanh	Hóa Học-Thanh	Hóa Học-Thanh
Thứ 5	_	C Tiết 3	-	Vật Lí-Anh	Vật Lí-Anh
Thứ 5	_	C Tiết 4	-	Ngữ Văn-Hương	Ngữ Văn-Hương
Thứ 5	_	C Tiết 5	-	SHL	SHL
Thứ 6	_	C Tiết 1	Toán-Minh	Toán-Minh	-
Thứ 6	_	C Tiết 2	Sinh Học-Liên	Sinh Học-Liên	-
Thứ 6	_	C Tiết 3	Vật Lí-Anh	Vật Lí-Anh	-
Thứ 6	_	C Tiết 4	Vật Lí-Anh	Vật Lí-Anh	-
Thứ 6	_	C Tiết 5	-	-	-
Thứ 7	_	C Tiết 1	-	-	-
Thứ 7	_	C Tiết 2	-	-	-
Thứ 7	_	C Tiết 3	-	-	-
Thứ 7	_	C Tiết 4	-	-	-
Thứ 7	_	C Tiết 5	-	-	-
```

**Lưu ý**: Dùng Tab (`\t`) để phân tách cột khi copy-paste vào Excel.

---

## 🧪 Cách Test

1. Tạo file Excel theo cấu trúc trên
2. Lưu file với tên: `TKB_Sample.xlsx`
3. Mở ứng dụng TKB Converter
4. Upload file Excel
5. Chọn lớp từ dropdown (VD: Lớp 10A1)
6. Kiểm tra preview bảng TKB
7. Nhấn "Tải File Word" để xuất
8. Mở file Word và kiểm tra

✅ **Nếu tất cả hoạt động đúng, ứng dụng đã setup thành công!**

---

## 📝 Ghi Chú

- Nếu bạn có nhiều lớp hơn (10A1 đến 10A20), chỉ cần thêm cột và dữ liệu tương ứng
- Cấu trúc file phải nhất quán (cột A: Thứ, C: Tiết, D+: Tên lớp)
- Dữ liệu tên môn có thể linh hoạt, ứng dụng sẽ tự động làm sạch

---

**Happy Testing! 🎉**