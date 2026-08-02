// Challenger List Generator - JS Conversion of Challenger.bas (Import_Challenger_MultipleFiles)

class ChallengerGenerator {
    constructor() {
        this.records = []; // [{ seq, consignee, item, itemDesc, shipQua }]
        this.searchQuery = "";
        this.init();
    }

    init() {
        this.loadFromStorage();
        this.renderUI();
    }

    loadFromStorage() {
        const stored = localStorage.getItem("DO_Challenger_List_Data");
        if (stored) {
            try {
                const parsed = JSON.parse(stored);
                this.records = Array.isArray(parsed) ? parsed : [];
            } catch (e) {
                console.error("Failed to load Challenger List data", e);
                this.records = [];
            }
        }
    }

    saveToStorage() {
        localStorage.setItem("DO_Challenger_List_Data", JSON.stringify(this.records));
    }

    reset() {
        if (this.records.length > 0 && !confirm("Are you sure you want to reset the Challenger List?")) return;
        this.records = [];
        localStorage.removeItem("DO_Challenger_List_Data");
        const fileInput = document.getElementById("challengerFilePicker");
        if (fileInput) fileInput.value = "";
        this.renderUI();
    }

    // Main entry point for uploading SONY - ROUTE OUTBOUND source files
    async handleFileUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        for (const file of files) {
            await this.processSourceFile(file);
        }

        event.target.value = ""; // Reset file picker
    }

    async processSourceFile(file) {
        const fileName = file.name;
        // Validation check for source filename (matching Challenger.bas naming convention)
        if (!fileName.toUpperCase().includes("SONY - ROUTE OUTBOUND")) {
            const proceed = confirm(`[WARNING] Selected file "${fileName}" does not match expected naming convention ("SONY - ROUTE OUTBOUND").\n\nDo you want to process it anyway?`);
            if (!proceed) return;
        }

        const dataBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(dataBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

        if (rawRows.length < 2) {
            alert(`[ERROR] File "${fileName}" has no data rows!`);
            return;
        }

        // Locate header row & map column indexes (matching Challenger.bas fixed columns A, D, J, K, L)
        const colMap = {
            seq:  0,
            consignee: 3,
            shipQua: 9,
            item: 10,
            itemDesc: 11
        };
        let dataStartRow = 1; // Header at row 1, data starts at row 2 (index 1)

        for (let r = 0; r < Math.min(5, rawRows.length); r++) {
            const rowText = (rawRows[r] || []).map(cell => String(cell || "").toLowerCase().trim());
            const sIdx = rowText.findIndex(t => t === "seq" || t === "sequence" || t === "seq no");
            const cIdx = rowText.findIndex(t => t === "consignee" || t === "customer name");
            const qIdx = rowText.findIndex(t => t === "ship_qua" || t === "ship qty" || t === "quantity");
            const iIdx = rowText.findIndex(t => t === "item" || t === "item code" || t === "product code");
            const dIdx = rowText.findIndex(t => t === "item_desc" || t === "item description" || t === "description");

            if (sIdx !== -1) colMap.seq = sIdx;
            if (cIdx !== -1) colMap.consignee = cIdx;
            if (qIdx !== -1) colMap.shipQua = qIdx;
            if (iIdx !== -1) colMap.item = iIdx;
            if (dIdx !== -1) colMap.itemDesc = dIdx;

            if (sIdx !== -1 || cIdx !== -1 || qIdx !== -1 || iIdx !== -1 || dIdx !== -1) {
                dataStartRow = r + 1;
                break;
            }
        }

        const newRecords = [];
        const skippedRecords = [];

        // Loop rows (matching: For i = 2 To LastRow)
        for (let i = dataStartRow; i < rawRows.length; i++) {
            const row = rawRows[i];
            if (!row || row.length === 0) continue;

            const consigneeVal = String(row[colMap.consignee] || "").trim();
            if (!consigneeVal) continue;

            // Include only rows where Consignee contains "Challenger" (matching InStr on Column D)
            if (!consigneeVal.toLowerCase().includes("challenger")) {
                skippedRecords.push(consigneeVal);
                continue;
            }

            newRecords.push({
                seq: String(row[colMap.seq] ?? "").trim(),
                consignee: consigneeVal,
                item: String(row[colMap.item] ?? "").trim(),
                itemDesc: String(row[colMap.itemDesc] ?? "").trim(),
                shipQua: parseInt(row[colMap.shipQua]) || 0
            });
        }

        if (newRecords.length === 0) {
            alert(`[WARNING] No Challenger DOs found in "${fileName}" (${skippedRecords.length} non-Challenger consignee row(s) skipped).`);
            return;
        }

        this.records = this.records.concat(newRecords);
        this.saveToStorage();
        this.renderUI();

        alert(`[SUCCESS] Extracted ${newRecords.length} Challenger line(s) from "${fileName}"! Total lines: ${this.records.length}.`);
    }

    setSearchQuery(val) {
        this.searchQuery = String(val || "").trim().toLowerCase();
        this.renderTable();
    }

    renderUI() {
        this.renderKPIs();
        this.renderTable();
    }

    renderKPIs() {
        const uniqueSeq = new Set(this.records.map(r => r.seq).filter(Boolean));
        const totalQty = this.records.reduce((sum, r) => sum + r.shipQua, 0);
        const uniqueItems = new Set(this.records.map(r => r.item).filter(Boolean));

        const elDO = document.getElementById("ch-kpi-do");
        const elLines = document.getElementById("ch-kpi-lines");
        const elQty = document.getElementById("ch-kpi-qty");
        const elSku = document.getElementById("ch-kpi-sku");

        if (elDO) elDO.innerText = uniqueSeq.size;
        if (elLines) elLines.innerText = this.records.length;
        if (elQty) elQty.innerText = totalQty.toLocaleString();
        if (elSku) elSku.innerText = uniqueItems.size;
    }

    renderTable() {
        const tbody = document.getElementById("challengerTableBody");
        if (!tbody) return;

        const elSearchCount = document.getElementById("challengerSearchCount");

        if (this.records.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:24px; color:#71717a;">No data loaded. Please upload SONY - ROUTE OUTBOUND source files.</td></tr>';
            if (elSearchCount) elSearchCount.style.display = "none";
            return;
        }

        const filteredRecords = this.records.filter(r => {
            if (!this.searchQuery) return true;
            const searchStr = `${r.seq} ${r.consignee} ${r.item} ${r.itemDesc}`.toLowerCase();
            return searchStr.includes(this.searchQuery);
        });

        if (elSearchCount) {
            elSearchCount.style.display = "inline-block";
            if (this.searchQuery) {
                elSearchCount.innerText = `${filteredRecords.length} / ${this.records.length} rows`;
            } else {
                elSearchCount.innerText = `${this.records.length} rows`;
            }
        }

        if (filteredRecords.length === 0) {
            tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:24px; color:#71717a;">No Challenger rows match search "${this.searchQuery}".</td></tr>`;
            return;
        }

        let html = "";
        filteredRecords.forEach((r, idx) => {
            html += `<tr>
                <td class="challenger-seq">${r.seq || '-'}</td>
                <td class="challenger-consignee">${r.consignee}</td>
                <td class="challenger-item">${r.item || '-'}</td>
                <td class="challenger-desc">${r.itemDesc || '-'}</td>
                <td class="challenger-qty">${r.shipQua.toLocaleString()}</td>
            </tr>`;
        });

        tbody.innerHTML = html;
    }

    // Export Workbook matching Challenger.bas output formatting
    async exportToExcel() {
        if (this.records.length === 0) {
            alert("[WARNING] No Challenger data to export! Please upload source files first.");
            return;
        }

        const workbook = XLSX.utils.book_new();

        // Style presets matching Challenger.bas
        const styleHeader = {
            font: { name: "Aptos Display", sz: 12, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "2F5597" } }, // RGB(47,85,151)
            alignment: { horizontal: "center", vertical: "center" }
        };

        const styleNormalData = { font: { name: "Aptos Display", sz: 11, bold: false } };

        const styleZebra = { font: { name: "Aptos Display", sz: 11, bold: false }, fill: { fgColor: { rgb: "F2F2F2" } } }; // RGB(242,242,242)

        const styleTitle = {
            font: { name: "Aptos Display", sz: 11, bold: true }
        };

        const sheetData = [];

        // Row 1: Challenger DO count (unique SEQ) - matching dict.Count
        const uniqueSeq = new Set(this.records.map(r => r.seq).filter(Boolean));
        sheetData.push(["Challenger DO:", uniqueSeq.size]);

        // Row 2: Headers
        sheetData.push(["SEQ", "CONSIGNEE", "ITEM", "ITEM_DESC", "SHIP_QUA"]);

        // Row 3+: Data Rows
        this.records.forEach(r => {
            sheetData.push([r.seq, r.consignee, r.item, r.itemDesc, r.shipQua]);
        });

        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        const lastR = this.records.length + 2;

        // Apply Row 1 Title Style (bold)
        ws['A1'].s = styleTitle;
        ws['B1'].s = styleTitle;

        // Apply Row 2 Header Style (blue fill, white bold, centered)
        ['A', 'B', 'C', 'D', 'E'].forEach(col => {
            const cellRef = `${col}2`;
            if (ws[cellRef]) ws[cellRef].s = styleHeader;
        });

        // Apply Zebra striping to Data Rows (Row 3+) - matching flipped parity from BAS
        this.records.forEach((rec, rIdx) => {
            const rowNum = rIdx + 3;
            const isZebra = (rowNum % 2 === 1);
            const style = isZebra ? styleZebra : styleNormalData;
            ['A', 'B', 'C', 'D', 'E'].forEach(col => {
                const cellRef = `${col}${rowNum}`;
                if (ws[cellRef]) ws[cellRef].s = style;
            });
        });

        // Borders (thin gray) covering header + data rows (Row 2 to lastR)
        const borderStyle = {
            top: { style: "thin", color: { rgb: "C8C8C8" } },
            bottom: { style: "thin", color: { rgb: "C8C8C8" } },
            left: { style: "thin", color: { rgb: "C8C8C8" } },
            right: { style: "thin", color: { rgb: "C8C8C8" } }
        };

        for (let r = 2; r <= lastR; r++) {
            ['A', 'B', 'C', 'D', 'E'].forEach(col => {
                const cellRef = `${col}${r}`;
                if (!ws[cellRef]) ws[cellRef] = { v: "" };
                if (!ws[cellRef].s) ws[cellRef].s = {};
                ws[cellRef].s.border = borderStyle;
            });
        }

        // Column Widths (AutoFit approximation matching BAS Columns.AutoFit)
        const seqWidth = Math.max(...this.records.map(r => String(r.seq).length), "SEQ".length);
        const consigneeWidth = Math.max(...this.records.map(r => String(r.consignee).length), "CONSIGNEE".length);
        const itemWidth = Math.max(...this.records.map(r => String(r.item).length), "ITEM".length);
        const itemDescWidth = Math.max(...this.records.map(r => String(r.itemDesc).length), "ITEM_DESC".length);
        const shipQuaWidth = Math.max(...this.records.map(r => String(r.shipQua).length), "SHIP_QUA".length);

        ws['!cols'] = [
            { wch: Math.min(seqWidth + 2, 20) },
            { wch: Math.min(consigneeWidth + 2, 40) },
            { wch: Math.min(itemWidth + 2, 25) },
            { wch: Math.min(itemDescWidth + 2, 40) },
            { wch: Math.min(shipQuaWidth + 2, 14) }
        ];

        XLSX.utils.book_append_sheet(workbook, ws, "Challenger List");

        // Determine export file name (matching: DO_Challenger_List ddmmyyyy.xlsx with tomorrow's date)
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        const dd = String(tomorrow.getDate()).padStart(2, '0');
        const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
        const yyyy = tomorrow.getFullYear();
        const exportFileName = `DO_Challenger_List ${dd}${mm}${yyyy}.xlsx`;

        // Explicit HTML5 Blob download (bypasses SheetJS internal write constraints on file:///)
        try {
            const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);

            const a = document.createElement("a");
            a.style.display = "none";
            a.href = url;
            a.download = exportFileName;

            document.body.appendChild(a);
            a.click();

            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 500);
        } catch (err) {
            console.error("Blob export failed:", err);
            alert("[ERROR] Could not save the file. Check browser permissions.");
        }
    }
}

// Global Instance
let challengerGenerator = null;
document.addEventListener("DOMContentLoaded", () => {
    challengerGenerator = new ChallengerGenerator();
});
