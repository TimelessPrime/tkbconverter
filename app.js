class TKBConverter {
    constructor() {
        this.workbook = null;
        this.worksheet = null;
        this.classes = [];
        this.selectedClass = null;
        this.classData = null;
        this.timeSlots = {
            morning: ['Tiết 1', 'Tiết 2', 'Tiết 3', 'Tiết 4', 'Tiết 5'],
            afternoon: ['Tiết 1', 'Tiết 2', 'Tiết 3', 'Tiết 4', 'Tiết 5']
        };
        this.days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
        this.initEventListeners();
    }

    initEventListeners() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const classDropdown = document.getElementById('classDropdown');
        const exportBtn = document.getElementById('exportBtn');

        // Upload events
        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        uploadArea.addEventListener('drop', (e) => this.handleFileDrop(e));
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Class selection
        classDropdown.addEventListener('change', (e) => this.handleClassSelect(e));

        // Export
        exportBtn.addEventListener('click', () => this.exportToWord());
    }

    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        document.getElementById('uploadArea').classList.add('dragover');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        document.getElementById('uploadArea').classList.remove('dragover');
    }

    handleFileDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        document.getElementById('uploadArea').classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    processFile(file) {
        const statusDiv = document.getElementById('fileStatus');
        statusDiv.className = 'file-status info';
        statusDiv.textContent = '⏳ Đang xử lý file...';

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });
                this.worksheet = this.workbook.Sheets[this.workbook.SheetNames[0]];

                this.extractClasses();
                statusDiv.className = 'file-status success';
                statusDiv.textContent = `✅ Tải file thành công! (Tìm thấy ${this.classes.length} lớp)`;

                this.showClassSection();
            } catch (error) {
                statusDiv.className = 'file-status error';
                statusDiv.textContent = `❌ Lỗi: ${error.message}`;
                console.error('File processing error:', error);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    extractClasses() {
        this.classes = [];
        const range = XLSX.utils.decode_range(this.worksheet['!ref']);

        // Quét tất cả ô để tìm tên lớp
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = this.worksheet[cellAddress];
                if (cell && cell.v) {
                    const value = String(cell.v).trim();
                    // Kiểm tra format lớp: 10A1 đến 10A14
                    if (/^10A\d+$/.test(value) && !this.classes.includes(value)) {
                        this.classes.push(value);
                    }
                }
            }
        }

        // Sắp xếp lớp theo thứ tự
        this.classes.sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });

        // Cập nhật dropdown
        const dropdown = document.getElementById('classDropdown');
        dropdown.innerHTML = '<option value="">-- Vui lòng chọn lớp --</option>';
        this.classes.forEach(cls => {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = `Lớp ${cls}`;
            dropdown.appendChild(option);
        });
    }

    findClassColumn(className) {
        const range = XLSX.utils.decode_range(this.worksheet['!ref']);
        
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = this.worksheet[cellAddress];
                if (cell && String(cell.v).trim() === className) {
                    return { column: col, startRow: row };
                }
            }
        }
        return null;
    }

    extractClassData(className) {
        const location = this.findClassColumn(className);
        if (!location) return null;

        const { column, startRow } = location;
        const schedule = {
            morning: {},
            afternoon: {}
        };

        // Khởi tạo cấu trúc
        this.days.forEach(day => {
            schedule.morning[day] = ['', '', '', '', ''];
            schedule.afternoon[day] = ['', '', '', '', ''];
        });

        // Dò từ dòng tiếp theo sau tên lớp
        const range = XLSX.utils.decode_range(this.worksheet['!ref']);
        let currentSession = 'morning';
        let foundAfternoon = false;

        for (let row = startRow + 1; row <= range.e.r; row++) {
            const dayCell = this.worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
            const timeCell = this.worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })];
            const subjectCell = this.worksheet[XLSX.utils.encode_cell({ r: row, c: column })];

            if (!dayCell || !timeCell || !subjectCell) continue;

            const day = String(dayCell.v).trim();
            const time = String(timeCell.v).trim();
            const subject = String(subjectCell.v).trim();

            // Kiểm tra buổi chiều
            if (time.includes('C') || (foundAfternoon && day.startsWith('Thứ'))) {
                currentSession = 'afternoon';
                foundAfternoon = true;
            } else if (day.startsWith('Thứ')) {
                // Nếu gặp "Thứ" và chưa gặp "C", vẫn là buổi sáng
                if (!time.includes('C')) currentSession = 'morning';
            }

            // Kiểm tra xem có phải ngày học không
            if (this.days.includes(day)) {
                const timeMatch = time.match(/\d+/);
                if (timeMatch) {
                    const timeIndex = parseInt(timeMatch[0]) - 1;
                    if (timeIndex >= 0 && timeIndex < 5) {
                        const cleanSubject = this.cleanSubjectName(subject);
                        schedule[currentSession][day][timeIndex] = cleanSubject;
                    }
                }
            }

            // Dừng khi gặp lớp khác
            if (day.match(/^10A\d+$/) && day !== className) {
                break;
            }
        }

        return schedule;
    }

    cleanSubjectName(subject) {
        if (!subject) return '';

        let cleaned = subject;

        // Loại bỏ tên giáo viên (phần sau dấu - hoặc phần cuối)
        if (cleaned.includes('-')) {
            cleaned = cleaned.split('-')[0].trim();
        }

        // Chuẩn hóa tên môn học
        const replacements = {
            'Chào cờ': 'SHDC',
            'HĐTNHN': 'HĐ TNHN',
            'HĐ TNHN': 'HĐ TNHN',
            'ND GDĐP': 'ND GDĐP',
            'SHL': 'SHL',
            'GDTC': 'GDTC',
            'SHDC': 'SHDC'
        };

        for (let [original, replacement] of Object.entries(replacements)) {
            if (cleaned.includes(original)) {
                cleaned = cleaned.replace(original, replacement);
            }
        }

        return cleaned.trim();
    }

    handleClassSelect(e) {
        this.selectedClass = e.target.value;
        if (!this.selectedClass) {
            document.getElementById('previewSection').style.display = 'none';
            document.getElementById('exportSection').style.display = 'none';
            return;
        }

        const statusDiv = document.getElementById('classStatus');
        statusDiv.className = 'status-message show info';
        statusDiv.textContent = '⏳ Đang xử lý thời khóa biểu...';

        // Xử lý bất đồng bộ
        setTimeout(() => {
            try {
                this.classData = this.extractClassData(this.selectedClass);
                if (this.classData) {
                    this.displayPreview();
                    statusDiv.className = 'status-message show success';
                    statusDiv.textContent = '✅ Thời khóa biểu sẵn sàng!';
                    document.getElementById('previewSection').style.display = 'block';
                    document.getElementById('exportSection').style.display = 'block';
                } else {
                    throw new Error('Không tìm thấy dữ liệu lớp này');
                }
            } catch (error) {
                statusDiv.className = 'status-message show error';
                statusDiv.textContent = `❌ Lỗi: ${error.message}`;
                console.error('Class processing error:', error);
            }
        }, 100);
    }

    displayPreview() {
        const previewDiv = document.getElementById('previewTable');
        let html = `<h3 style="margin-bottom: 15px;">Buổi Sáng</h3>`;
        html += this.generateTableHTML(this.classData.morning);
        html += `<h3 style="margin: 25px 0 15px 0;">Buổi Chiều</h3>`;
        html += this.generateTableHTML(this.classData.afternoon);
        previewDiv.innerHTML = html;
    }

    generateTableHTML(schedule) {
        let html = '<table class="preview-table"><thead><tr><th>Buổi</th><th>Tiết</th>';
        this.days.forEach(day => html += `<th>${day}</th>`);
        html += '</tr></thead><tbody>';

        this.timeSlots.morning.forEach((time, index) => {
            html += `<tr><td>S</td><td>${time}</td>`;
            this.days.forEach(day => {
                const subject = schedule[day] ? schedule[day][index] : '';
                html += `<td>${subject}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table>';
        return html;
    }

    async exportToWord() {
        const exportBtn = document.getElementById('exportBtn');
        const statusDiv = document.getElementById('exportStatus');
        
        exportBtn.disabled = true;
        statusDiv.className = 'status-message show info';
        statusDiv.textContent = '⏳ Đang tạo file Word...';

        try {
            const sections = [];

            // Tạo bảng buổi sáng
            sections.push(
                new docx.Paragraph({
                    text: `THỜI KHÓA BIỂU LỚP ${this.selectedClass} - BUỔI SÁNG`,
                    bold: true,
                    size: 24,
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 200 }
                })
            );

            sections.push(this.createScheduleTable(this.classData.morning));

            sections.push(
                new docx.Paragraph({
                    text: `\nTHỜI KHÓA BIỂU LỚP ${this.selectedClass} - BUỔI CHIỀU`,
                    bold: true,
                    size: 24,
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { before: 400, after: 200 }
                })
            );

            sections.push(this.createScheduleTable(this.classData.afternoon));

            // Tạo document
            const doc = new docx.Document({
                sections: [{
                    properties: {},
                    children: sections
                }]
            });

            // Lưu file
            const fileName = `TKB_Lop_${this.selectedClass}.docx`;
            docx.Packer.toBlob(doc).then(blob => {
                saveAs(blob, fileName);
                exportBtn.disabled = false;
                statusDiv.className = 'status-message show success';
                statusDiv.textContent = `✅ Đã tải file ${fileName}!`;
            });

        } catch (error) {
            exportBtn.disabled = false;
            statusDiv.className = 'status-message show error';
            statusDiv.textContent = `❌ Lỗi: ${error.message}`;
            console.error('Export error:', error);
        }
    }

    createScheduleTable(schedule) {
        const rows = [];

        // Header row
        const headerCells = [
            new docx.TableCell({
                children: [new docx.Paragraph({ text: 'Tiết', bold: true })],
                shading: { fill: 'D3D3D3' }
            })
        ];

        this.days.forEach(day => {
            headerCells.push(
                new docx.TableCell({
                    children: [new docx.Paragraph({ text: day, bold: true })],
                    shading: { fill: 'D3D3D3' }
                })
            );
        });

        rows.push(new docx.TableRow({ children: headerCells }));

        // Data rows
        this.timeSlots.morning.forEach((time, index) => {
            const cells = [
                new docx.TableCell({
                    children: [new docx.Paragraph({ text: time, bold: true })],
                    shading: { fill: 'E8E8E8' }
                })
            ];

            this.days.forEach(day => {
                const subject = (schedule[day] && schedule[day][index]) ? schedule[day][index] : '';
                cells.push(
                    new docx.TableCell({
                        children: [new docx.Paragraph({ text: subject })]
                    })
                );
            });

            rows.push(new docx.TableRow({ children: cells }));
        });

        return new docx.Table({
            rows: rows,
            width: {
                size: 100,
                type: docx.WidthType.PERCENTAGE
            },
            borders: {
                top: { style: docx.BorderStyle.SINGLE, size: 6, color: '000000' },
                bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: '000000' },
                left: { style: docx.BorderStyle.SINGLE, size: 6, color: '000000' },
                right: { style: docx.BorderStyle.SINGLE, size: 6, color: '000000' },
                insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 6, color: '000000' },
                insideVertical: { style: docx.BorderStyle.SINGLE, size: 6, color: '000000' }
            }
        });
    }

    showClassSection() {
        document.getElementById('classSection').style.display = 'block';
    }
}

// Khởi tạo ứng dụng khi tài liệu sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
    new TKBConverter();
});