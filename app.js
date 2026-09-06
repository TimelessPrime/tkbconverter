class TKBConverter {
    constructor() {
        this.workbook = null;
        this.worksheet = null;
        this.classes = [];
        this.selectedClass = null;
        this.classData = null;
        this.days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];
        this.isDebug = false;
        this.fileName = '';
        this.appliedDateStr = '05/09/2026'; // Ngày mặc định nếu không tìm thấy trong tên file
        
        this.initDebugMode();
        this.initEventListeners();
    }

    initDebugMode() {
        const toggleBtn = document.getElementById('toggleDebugBtn');
        const consoleDiv = document.getElementById('debugConsole');

        if (!toggleBtn || !consoleDiv) return;

        const originalLog = console.log;
        const originalError = console.error;

        const appendLog = (type, msg) => {
            if (!this.isDebug) return;
            const time = new Date().toLocaleTimeString();
            const color = type === 'ERROR' ? '#fca5a5' : '#a7f3d0';
            consoleDiv.innerHTML += `<div style="color: ${color}; margin-bottom: 2px;">[${time}] [${type}] ${msg}</div>`;
            consoleDiv.scrollTop = consoleDiv.scrollHeight;
        };

        console.log = (...args) => {
            originalLog.apply(console, args);
            appendLog('INFO', args.join(' '));
        };

        console.error = (...args) => {
            originalError.apply(console, args);
            appendLog('ERROR', args.join(' '));
        };

        toggleBtn.onclick = (e) => {
            if (e) {
                e.preventDefault();
                e.stopPropagation();
            }

            this.isDebug = !this.isDebug;
            consoleDiv.style.display = this.isDebug ? 'block' : 'none';
            toggleBtn.style.background = this.isDebug ? 'rgba(56, 189, 248, 0.25)' : 'rgba(255,255,255,0.1)';
            toggleBtn.style.color = this.isDebug ? '#38bdf8' : '#cbd5e1';
            toggleBtn.textContent = this.isDebug ? '🐛 Disable Debug' : '🐛 Enable Debug';
            
            if (this.isDebug) {
                console.log('Chế độ Debug đã được bật.');
            }
        };
    }

    initEventListeners() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const loadUrlBtn = document.getElementById('loadUrlBtn');
        const classDropdown = document.getElementById('classDropdown');
        const exportBtn = document.getElementById('exportBtn');

        if (uploadArea && fileInput) {
            uploadArea.addEventListener('click', (e) => {
                if (e.target !== fileInput) {
                    fileInput.click();
                }
            });

            uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
            uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
            uploadArea.addEventListener('drop', (e) => this.handleFileDrop(e));
            fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        }

        if (loadUrlBtn) {
            loadUrlBtn.addEventListener('click', () => this.fetchFileFromUrl());
        }

        if (classDropdown) {
            classDropdown.addEventListener('change', (e) => this.handleClassSelect(e));
        }

        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.exportToExcel());
        }
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
        if (e.dataTransfer && e.dataTransfer.files.length > 0) {
            this.processFile(e.dataTransfer.files[0]);
        }
    }

    handleFileSelect(e) {
        if (e.target.files && e.target.files.length > 0) {
            this.processFile(e.target.files[0]);
        }
    }

    async fetchFileFromUrl() {
        const urlInput = document.getElementById('fileUrlInput');
        const statusDiv = document.getElementById('fileStatus');
        const url = urlInput ? urlInput.value.trim() : '';

        if (!url) {
            statusDiv.className = 'file-status error';
            statusDiv.textContent = '🔴 Vui lòng nhập link URL hợp lệ!';
            return;
        }

        statusDiv.className = 'file-status info';
        statusDiv.textContent = '🔵 Đang tải file từ URL...';

        try {
            const response = await fetch(url);
            if (!response.ok) throw new Error(`HTTP Error Status: ${response.status}`);

            const blob = await response.blob();
            if (blob.size === 0) throw new Error('File tải về rỗng (0 bytes)');

            const fileName = url.split('/').pop().split('?')[0] || 'TKB_Remote.xlsx';
            const file = new File([blob], fileName, { type: blob.type });

            this.processFile(file);
        } catch (err) {
            statusDiv.className = 'file-status error';
            statusDiv.textContent = `🔴 Upload không thành công. Lỗi kết nối URL hoặc CORS: ${err.message}`;
        }
    }

    // Trích xuất ngày áp dụng từ tên file (Ví dụ tìm chuỗi: 05.09.2026 hoặc 05-09-2026)
    extractDateFromFileName(fileName) {
        this.fileName = fileName;
        const dateMatch = fileName.match(/(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{4})/);
        if (dateMatch) {
            const day = dateMatch[1].padStart(2, '0');
            const month = dateMatch[2].padStart(2, '0');
            const year = dateMatch[3];
            this.appliedDateStr = `${day}/${month}/${year}`;
        } else {
            this.appliedDateStr = '05/09/2026';
        }
    }

    processFile(file) {
        const statusDiv = document.getElementById('fileStatus');

        if (!file || file.size === 0) {
            statusDiv.className = 'file-status error';
            statusDiv.textContent = '🔴 Upload không thành công (File rỗng - 0 bytes)!';
            return;
        }

        this.extractDateFromFileName(file.name);

        statusDiv.className = 'file-status info';
        statusDiv.textContent = `🔵 Đang đọc file: ${file.name}...`;

        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });
                this.worksheet = this.workbook.Sheets[this.workbook.SheetNames[0]];

                this.extractClasses();

                if (this.classes.length > 0) {
                    statusDiv.className = 'file-status success';
                    statusDiv.textContent = `🟢 Xử lý thành công! Nhận diện ${this.classes.length} lớp (Ngày áp dụng: ${this.appliedDateStr}).`;
                    document.getElementById('classSection').style.display = 'block';
                } else {
                    statusDiv.className = 'file-status error';
                    statusDiv.textContent = '🔴 Không tìm thấy dữ liệu lớp học phù hợp!';
                }

            } catch (error) {
                statusDiv.className = 'file-status error';
                statusDiv.textContent = `🔴 Xử lý không thành công: ${error.message}`;
            }
        };

        reader.onerror = () => {
            statusDiv.className = 'file-status error';
            statusDiv.textContent = '🔴 Upload không thành công. Lỗi đọc file!';
        };

        reader.readAsArrayBuffer(file);
    }

    getCellValue(r, c) {
        if (!this.worksheet) return '';
        
        const cellAddr = XLSX.utils.encode_cell({ r, c });
        let cell = this.worksheet[cellAddr];

        if (cell && cell.v !== undefined && cell.v !== null) {
            return String(cell.v).trim();
        }

        if (this.worksheet['!merges']) {
            for (let merge of this.worksheet['!merges']) {
                if (r >= merge.s.r && r <= merge.e.r && c >= merge.s.c && c <= merge.e.c) {
                    const masterCellAddr = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
                    const masterCell = this.worksheet[masterCellAddr];
                    if (masterCell && masterCell.v !== undefined) {
                        return String(masterCell.v).trim();
                    }
                }
            }
        }

        return '';
    }

    extractClasses() {
        this.classes = [];
        if (!this.worksheet || !this.worksheet['!ref']) return;

        const range = XLSX.utils.decode_range(this.worksheet['!ref']);

        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const val = this.getCellValue(r, c);
                if (/^10A\d+$/i.test(val) && !this.classes.includes(val)) {
                    this.classes.push(val);
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
                const cellAddr = XLSX.utils.encode_cell({ r, c });
                const cell = this.worksheet[cellAddr];
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

        const schedule = { afternoon: {} };
        this.days.forEach(d => {
            schedule.afternoon[d] = ['', '', '', '', ''];
        });

        const range = XLSX.utils.decode_range(this.worksheet['!ref']);

        for (let r = loc.row + 1; r <= range.e.r; r++) {
            const dayVal = this.getCellValue(r, 0); 
            
            let currentDay = null;
            if (dayVal) {
                currentDay = this.days.find(d => 
                    dayVal.toLowerCase().includes(d.toLowerCase()) || 
                    dayVal === d.replace('Thứ ', '').trim()
                );
            }

            if (!currentDay) continue;

            const tietVal = this.getCellValue(r, 2);
            const match = tietVal.match(/\d+/);
            if (!match) continue;

            const slotIdx = parseInt(match[0]) - 1;
            if (slotIdx < 0 || slotIdx > 4) continue;

            const subjectVal = this.getCellValue(r, loc.col);
            if (subjectVal) {
                schedule.afternoon[currentDay][slotIdx] = this.cleanSubject(subjectVal);
            }

            if (currentDay === 'Thứ 6' && slotIdx === 4) {
                break;
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
        statusDiv.textContent = '🔵 Đang xử lý trích xuất dữ liệu lớp...';

        setTimeout(() => {
            this.classData = this.extractClassData(this.selectedClass);
            if (this.classData) {
                this.renderPreview();
                statusDiv.className = 'status-message show success';
                statusDiv.textContent = '🟢 Xử lý thành công dữ liệu lớp!';
                document.getElementById('previewSection').style.display = 'block';
                document.getElementById('exportSection').style.display = 'block';
            } else {
                statusDiv.className = 'status-message show error';
                statusDiv.textContent = '🔴 Không tìm thấy thời khóa biểu của lớp này!';
            }
        }, 50);
    }

    renderPreview() {
        const preview = document.getElementById('previewTable');
        preview.innerHTML = `
            <h3 style="margin: 10px 0; color: #38bdf8; text-align: center;">🌆 Thời Khóa Biểu Buổi Chiều - Lớp ${this.selectedClass}</h3>
            <div class="preview-table-wrapper">
                ${this.buildHTMLTable(this.classData.afternoon, 'C')}
            </div>
        `;
    }

    buildHTMLTable(data, sessionLabel) {
        let html = `<table class="preview-table"><thead><tr><th>Buổi</th><th>Tiết</th>`;
        this.days.forEach(d => html += `<th>${d}</th>`);
        html += `</tr></thead><tbody>`;

        for (let i = 0; i < 5; i++) {
            html += `<tr>`;
            if (i === 0) {
                html += `<td rowspan="5" style="vertical-align: middle; font-weight: bold;">${sessionLabel}</td>`;
            }
            html += `<td>${i + 1}</td>`;
            this.days.forEach(d => {
                html += `<td>${data[d][i] || ''}</td>`;
            });
            html += `</tr>`;
        }
        html += `</tbody></table>`;
        return html;
    }

    // Xuất Excel đúng chuẩn hình mẫu
    exportToExcel() {
        const exportBtn = document.getElementById('exportBtn');
        const statusDiv = document.getElementById('exportStatus');

        if (typeof XLSX === 'undefined') {
            statusDiv.className = 'status-message show error';
            statusDiv.textContent = '🔴 Thư viện Excel chưa sẵn sàng. Vui lòng thử tải lại trang!';
            return;
        }

        exportBtn.disabled = true;
        statusDiv.className = 'status-message show info';
        statusDiv.textContent = '🔵 Đang tạo file Excel chuẩn khung...';

        try {
            // Định dạng tên lớp theo mẫu (10A9 -> 10/9)
            const classNumber = this.selectedClass.replace(/^10A/i, '');
            const titleClassStr = `10/${classNumber}`;

            const headerTitle = `THỜI KHÓA BIỂU LỚP ${titleClassStr} NĂM HỌC 2026-2027 ÁP DỤNG NGÀY ${this.appliedDateStr}`;

            // Cấu trúc dữ liệu mảng 2 chiều
            const excelData = [
                [headerTitle, "", "", "", "", "", ""], // Row 0: Dòng tiêu đề lớn
                ["Buổi", "Tiết", "Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"] // Row 1: Header bảng
            ];

            // Thêm 5 dòng tiết cho buổi chiều (C)
            for (let i = 0; i < 5; i++) {
                excelData.push([
                    i === 0 ? "C" : "", // Dòng đầu chứa "C", các dòng sau để trống để chuẩn bị Merge
                    i + 1,
                    ...this.days.map(d => this.classData.afternoon[d][i] || "")
                ]);
            }

            const ws = XLSX.utils.aoa_to_sheet(excelData);

            // Cấu hình Merged Cells (Gộp ô):
            // 1. Dòng 0: Gộp Cột 0 (A) -> Cột 6 (G) làm tiêu đề chính
            // 2. Cột 0 (Buổi): Gộp Dòng 2 (A3) -> Dòng 6 (A7) cho chữ C
            ws['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
                { s: { r: 2, c: 0 }, e: { r: 6, c: 0 } }
            ];

            // Thiết lập độ rộng cột
            ws['!cols'] = [
                { wch: 8 },  // Buổi
                { wch: 8 },  // Tiết
                { wch: 16 }, // Thứ 2
                { wch: 16 }, // Thứ 3
                { wch: 16 }, // Thứ 4
                { wch: 16 }, // Thứ 5
                { wch: 16 }  // Thứ 6
            ];

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, `TKB_Lop_${this.selectedClass}`);

            XLSX.writeFile(wb, `TKB_Lop_${this.selectedClass}.xlsx`);

            exportBtn.disabled = false;
            statusDiv.className = 'status-message show success';
            statusDiv.textContent = '🟢 Xử lý thành công! File Excel đã được tải về.';
        } catch (err) {
            exportBtn.disabled = false;
            statusDiv.className = 'status-message show error';
            statusDiv.textContent = `🔴 Xử lý không thành công: ${err.message}`;
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new TKBConverter();
});