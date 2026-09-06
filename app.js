class TKBConverter {
    constructor() {
        this.workbook = null;
        this.worksheet = null;
        this.classes = [];
        this.selectedClass = null;
        this.classData = null;
        this.days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
        this.initEventListeners();
    }

    initEventListeners() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const classDropdown = document.getElementById('classDropdown');
        const exportBtn = document.getElementById('exportBtn');

        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        uploadArea.addEventListener('drop', (e) => this.handleFileDrop(e));
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        classDropdown.addEventListener('change', (e) => this.handleClassSelect(e));
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
        if (e.dataTransfer.files.length > 0) {
            this.processFile(e.dataTransfer.files[0]);
        }
    }

    handleFileSelect(e) {
        if (e.target.files.length > 0) {
            this.processFile(e.target.files[0]);
        }
    }

    processFile(file) {
        const statusDiv = document.getElementById('fileStatus');
        statusDiv.className = 'file-status info';
        statusDiv.textContent = '⏳ Đang đọc file...';

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });
                this.worksheet = this.workbook.Sheets[this.workbook.SheetNames[0]];

                this.extractClasses();
                statusDiv.className = 'file-status success';
                statusDiv.textContent = `✅ Đã nhận diện ${this.classes.length} lớp.`;
                document.getElementById('classSection').style.display = 'block';
            } catch (error) {
                statusDiv.className = 'file-status error';
                statusDiv.textContent = `❌ Lỗi đọc file: ${error.message}`;
            }
        };
        reader.readAsArrayBuffer(file);
    }

    extractClasses() {
        this.classes = [];
        const range = XLSX.utils.decode_range(this.worksheet['!ref']);

        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = this.worksheet[XLSX.utils.encode_cell({ r, c })];
                if (cell && cell.v) {
                    const val = String(cell.v).trim();
                    if (/^10A\d+$/i.test(val) && !this.classes.includes(val)) {
                        this.classes.push(val);
                    }
                }
            }
        }

        this.classes.sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

        const dropdown = document.getElementById('classDropdown');
        dropdown.innerHTML = '<option value="">-- Chọn lớp cần xuất TKB --</option>';
        this.classes.forEach(cls => {
            const opt = document.createElement('option');
            opt.value = cls;
            opt.textContent = `Lớp ${cls}`;
            dropdown.appendChild(opt);
        });
    }

    findClassLocation(className) {
        const range = XLSX.utils.decode_range(this.worksheet['!ref']);
        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = this.worksheet[XLSX.utils.encode_cell({ r, c })];
                if (cell && String(cell.v).trim() === className) {
                    return { col: c, row: r };
                }
            }
        }
        return null;
    }

    extractClassData(className) {
        const loc = this.findClassLocation(className);
        if (!loc) return null;

        const schedule = { morning: {}, afternoon: {} };
        this.days.forEach(d => {
            schedule.morning[d] = ['', '', '', '', ''];
            schedule.afternoon[d] = ['', '', '', '', ''];
        });

        const range = XLSX.utils.decode_range(this.worksheet['!ref']);
        let lastDay = 'Thứ 2';

        for (let r = loc.row + 1; r <= Math.min(loc.row + 55, range.e.r); r++) {
            const dayCell = this.worksheet[XLSX.utils.encode_cell({ r, c: 0 })];
            const timeCell = this.worksheet[XLSX.utils.encode_cell({ r, c: 2 })];
            const subjectCell = this.worksheet[XLSX.utils.encode_cell({ r, c: loc.col })];

            if (dayCell && dayCell.v && this.days.includes(String(dayCell.v).trim())) {
                lastDay = String(dayCell.v).trim();
            }

            if (!timeCell || !timeCell.v) continue;

            const timeVal = String(timeCell.v).trim();
            const subjectVal = subjectCell && subjectCell.v ? String(subjectCell.v).trim() : '';

            const isAfternoon = r > loc.row + 25 || timeVal.includes('C');
            const session = isAfternoon ? 'afternoon' : 'morning';

            const matchSlot = timeVal.match(/\d+/);
            if (matchSlot) {
                const slotIdx = parseInt(matchSlot[0]) - 1;
                if (slotIdx >= 0 && slotIdx < 5) {
                    schedule[session][lastDay][slotIdx] = this.cleanSubject(subjectVal);
                }
            }
        }

        return schedule;
    }

    cleanSubject(raw) {
        if (!raw) return '';
        let s = raw.split('-')[0].trim();
        s = s.replace('Chào cờ', 'SHDC')
             .replace('HĐTNHN', 'HĐ TNHN')
             .replace('ND GDĐP', 'GDĐP');
        return s;
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
        statusDiv.textContent = '⏳ Đang trích xuất dữ liệu...';

        setTimeout(() => {
            this.classData = this.extractClassData(this.selectedClass);
            if (this.classData) {
                this.renderPreview();
                statusDiv.className = 'status-message show success';
                statusDiv.textContent = '✅ Đã trích xuất thành công!';
                document.getElementById('previewSection').style.display = 'block';
                document.getElementById('exportSection').style.display = 'block';
            } else {
                statusDiv.className = 'status-message show error';
                statusDiv.textContent = '❌ Không tìm thấy dữ liệu lớp.';
            }
        }, 50);
    }

    renderPreview() {
        const preview = document.getElementById('previewTable');
        preview.innerHTML = `
            <h3 style="margin-bottom:10px;">Sáng</h3>
            ${this.buildHTMLTable(this.classData.morning, 'S')}
            <h3 style="margin:20px 0 10px 0;">Chiều</h3>
            ${this.buildHTMLTable(this.classData.afternoon, 'C')}
        `;
    }

    buildHTMLTable(data, sessionLabel) {
        let html = `<table class="preview-table"><thead><tr><th>Buổi</th><th>Tiết</th>`;
        this.days.forEach(d => html += `<th>${d}</th>`);
        html += `</tr></thead><tbody>`;

        for (let i = 0; i < 5; i++) {
            html += `<tr><td>${sessionLabel}</td><td>${i + 1}</td>`;
            this.days.forEach(d => {
                html += `<td>${data[d][i] || ''}</td>`;
            });
            html += `</tr>`;
        }
        html += `</tbody></table>`;
        return html;
    }

    async exportToWord() {
        const exportBtn = document.getElementById('exportBtn');
        const statusDiv = document.getElementById('exportStatus');

        exportBtn.disabled = true;
        statusDiv.className = 'status-message show info';
        statusDiv.textContent = '⏳ Đang tạo file Word...';

        try {
            const docxLib = window.docx;
            const { Document, Packer, Paragraph, Table, TableRow, TableCell, AlignmentType, WidthType, BorderStyle } = docxLib;

            const makeTable = (data, sessionTitle) => {
                const rows = [];
                
                const headerCells = [
                    new TableCell({ children: [new Paragraph({ text: "Tiết", bold: true })], shading: { fill: "E0E0E0" } }),
                    ...this.days.map(d => new TableCell({ children: [new Paragraph({ text: d, bold: true })], shading: { fill: "E0E0E0" } }))
                ];
                rows.push(new TableRow({ children: headerCells }));

                for (let i = 0; i < 5; i++) {
                    const cells = [
                        new TableCell({ children: [new Paragraph({ text: `Tiết ${i + 1}`, bold: true })] }),
                        ...this.days.map(d => new TableCell({ children: [new Paragraph({ text: data[d][i] || "" })] }))
                    ];
                    rows.push(new TableRow({ children: cells }));
                }

                return [
                    new Paragraph({ text: sessionTitle, bold: true, size: 24, spacing: { before: 200, after: 100 } }),
                    new Table({
                        rows,
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 4 },
                            bottom: { style: BorderStyle.SINGLE, size: 4 },
                            left: { style: BorderStyle.SINGLE, size: 4 },
                            right: { style: BorderStyle.SINGLE, size: 4 },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 4 },
                            insideVertical: { style: BorderStyle.SINGLE, size: 4 }
                        }
                    })
                ];
            };

            const doc = new Document({
                sections: [{
                    children: [
                        new Paragraph({ text: `THỜI KHÓA BIỂU - LỚP ${this.selectedClass}`, bold: true, size: 32, alignment: AlignmentType.CENTER }),
                        ...makeTable(this.classData.morning, "BUỔI SÁNG"),
                        ...makeTable(this.classData.afternoon, "BUỔI CHIỀU")
                    ]
                }]
            });

            const blob = await Packer.toBlob(doc);
            saveAs(blob, `TKB_Lop_${this.selectedClass}.docx`);

            exportBtn.disabled = false;
            statusDiv.className = 'status-message show success';
            statusDiv.textContent = '✅ Đã tải file Word về máy!';
        } catch (err) {
            exportBtn.disabled = false;
            statusDiv.className = 'status-message show error';
            statusDiv.textContent = `❌ Lỗi xuất file: ${err.message}`;
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new TKBConverter();
});