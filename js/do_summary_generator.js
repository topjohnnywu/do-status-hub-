// DO Summary List Generator - JavaScript Conversion of BatachSummary.bas & FinalSummary.bas

class DOSummaryGenerator {
    constructor() {
        this.batches = []; // [{ batchName: 'Batch 01', waveNumber: '01', records: [...] }]
        this.currentBatchIndex = 0;
        this.hasCompiledFinalSummary = false;
        this.searchQuery = "";
        this.targetFileName = "";
        this.init();
    }

    init() {
        this.loadFromStorage();
        this.renderUI();
    }

    loadFromStorage() {
        const stored = localStorage.getItem("DO_Summary_Generator_Data");
        if (stored) {
            try {
                this.batches = JSON.parse(stored);
                // Ensure all records have selected property initialized
                this.batches.forEach(b => {
                    if (b.records) {
                        b.records.forEach(r => {
                            if (r.selected === undefined) r.selected = true;
                        });
                    }
                });
            } catch (e) {
                console.error("Failed to load DO Summary Generator data", e);
                this.batches = [];
            }
        }
        this.hasCompiledFinalSummary = localStorage.getItem("DO_Summary_Generator_Compiled") === "1";
        this.targetFileName = localStorage.getItem("DO_Summary_Generator_FileName") || "";
        this.generatorDate = localStorage.getItem("DO_Summary_Generator_Date") || "";
    }

    saveToStorage() {
        localStorage.setItem("DO_Summary_Generator_Data", JSON.stringify(this.batches));
        localStorage.setItem("DO_Summary_Generator_Compiled", this.hasCompiledFinalSummary ? "1" : "0");
        if (this.targetFileName) {
            localStorage.setItem("DO_Summary_Generator_FileName", this.targetFileName);
        }
        if (this.generatorDate) {
            localStorage.setItem("DO_Summary_Generator_Date", this.generatorDate);
        }
    }

    reset() {
        if (this.batches.length > 0 && !confirm("Are you sure you want to reset all generator batches?")) return;
        this.batches = [];
        this.currentBatchIndex = 0;
        this.hasCompiledFinalSummary = false;
        localStorage.removeItem("DO_Summary_Generator_Data");
        localStorage.removeItem("DO_Summary_Generator_Compiled");
        localStorage.removeItem("DO_Summary_Generator_FileName");
        localStorage.removeItem("DO_Summary_Generator_Date");
        const fileInput = document.getElementById("sourceFilePicker");
        if (fileInput) fileInput.value = "";
        this.renderUI();
    }

    // Main entry point for uploading SONY DO Summary List CSV/Excel source files
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
        // Validation check for source filename
        if (!fileName.toUpperCase().includes("SONY - DO SUMMARY LIST") && !fileName.toUpperCase().includes("SUMMARY")) {
            const proceed = confirm(`[WARNING] Selected file "${fileName}" does not match expected naming convention ("SONY - DO Summary List").\n\nDo you want to process it anyway?`);
            if (!proceed) return;
        }

        const dataBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(dataBuffer, { type: 'array' });

        // FIX 1: Sheet Target - Try finding "final summary", else FALLBACK TO SHEET 1 (Matches VBA: Set wsSource = wbSource.Sheets(1))
        let sheetName = workbook.SheetNames.find(name => {
            const lower = name.toLowerCase().trim();
            return lower.includes("final summary") || lower.includes("final");
        });

        if (!sheetName) {
            sheetName = workbook.SheetNames[0]; // Take 1st sheet automatically
        }

        const sheet = workbook.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

        if (!rawRows || rawRows.length < 2) {
            alert(`[ERROR] File "${fileName}" has no data rows!`);
            return;
        }

        // FIX 2: Header Map - Includes SONY CSV Aliases (SEQ, SHIP_TO, ZONE, CONSIGNEE_NAME1, TOTAL_ITEM)
        const headerRow = rawRows[0].map(h => String(h || "").trim().toUpperCase());
        
        const colMap = {
            invoiceNo: this.findColIdx(headerRow, ["INVOICE NO.", "INVOICE", "INVOICE_NO", "DO NO", "DO_NUMBER", "SEQ"]),
            division:  this.findColIdx(headerRow, ["DIVISION", "DIV"]),
            shpCode:   this.findColIdx(headerRow, ["SHP_CODE", "SHP CODE", "SHIP CODE", "SHIP_TO", "SHIP TO"]),
            route:     this.findColIdx(headerRow, ["ROUTE", "ZONE"]),
            consignee: this.findColIdx(headerRow, ["CONSIGNEE_NAME", "CONSIGNEE", "CUSTOMER", "CONSIGNEE_NAME1", "CONSIGNEE NAME1", "CONSIGNEE NAME"]),
            addr1:     this.findColIdx(headerRow, ["ADDRESS1", "ADDRESS 1", "ADDR1"]),
            addr2:     this.findColIdx(headerRow, ["ADDRESS2", "ADDRESS 2", "ADDR2"]),
            addr3:     this.findColIdx(headerRow, ["ADDRESS3", "ADDRESS 3", "ADDR3"]),
            volume:    this.findColIdx(headerRow, ["VOLUME", "M3", "VOLUME (M3)"]),
            qty:       this.findColIdx(headerRow, ["QTY", "QUANTITY"]),
            sku:       this.findColIdx(headerRow, ["TOTAL SKU", "SKU", "SKU COUNT", "TOTAL_ITEM", "TOTAL ITEM"]),
            remark:    this.findColIdx(headerRow, ["REMARK", "REMARKS"])
        };

        const parsedRecords = [];
        let missingRouteCount = 0;
        let missingVolumeCount = 0;

        for (let i = 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            if (!row || row.length === 0) continue;

            const invoiceNo = String(row[colMap.invoiceNo !== -1 ? colMap.invoiceNo : 0] || "").trim();
            if (!invoiceNo) continue;

            const routeVal = String(row[colMap.route !== -1 ? colMap.route : 3] || "").trim();
            const rawVolStr = String(row[colMap.volume !== -1 ? colMap.volume : 8] ?? "").trim();
            const cleanedVolStr = rawVolStr.toLowerCase().replace(/m3|m³/g, "").trim();
            const volVal = parseFloat(cleanedVolStr) || 0;
            
            const qtyVal = parseInt(row[colMap.qty !== -1 ? colMap.qty : 9]) || 0;
            const skuVal = parseInt(row[colMap.sku !== -1 ? colMap.sku : 10]) || 1;

            const missingRoute = !routeVal || routeVal === "-" || routeVal.toUpperCase() === "MISSING";
            const missingVolume = volVal <= 0 || cleanedVolStr === "" || cleanedVolStr === "-" || cleanedVolStr === "0";

            if (missingRoute) missingRouteCount++;
            if (missingVolume) missingVolumeCount++;

            parsedRecords.push({
                invoiceNo: invoiceNo,
                division: String(row[colMap.division !== -1 ? colMap.division : 1] || "").trim(),
                shpCode: String(row[colMap.shpCode !== -1 ? colMap.shpCode : 2] || "").trim(),
                route: routeVal,
                consignee: String(row[colMap.consignee !== -1 ? colMap.consignee : 4] || "").trim(),
                addr1: String(row[colMap.addr1 !== -1 ? colMap.addr1 : 5] || "").trim(),
                addr2: String(row[colMap.addr2 !== -1 ? colMap.addr2 : 6] || "").trim(),
                addr3: String(row[colMap.addr3 !== -1 ? colMap.addr3 : 7] || "").trim(),
                volume: volVal,
                qty: qtyVal,
                sku: skuVal,
                remark: "", // Ignore CSV summary text; leave blank for manual entry
                missingRoute: missingRoute,
                missingVolume: missingVolume,
                selected: true
            });
        }

        if (parsedRecords.length === 0) {
            alert(`[WARNING] No valid DO records found in "${fileName}"!`);
            return;
        }

        // Show Missing Information Alert if applicable
        if (missingRouteCount > 0 || missingVolumeCount > 0) {
            alert(`[NOTICE] Missing Information Detected in Source File:\n` +
                  (missingRouteCount > 0 ? `- ${missingRouteCount} DO(s) with empty ROUTE\n` : '') +
                  (missingVolumeCount > 0 ? `- ${missingVolumeCount} DO(s) with 0 or empty VOLUME (M3)\n` : '') +
                  `\nMissing values have been flagged in red for review.`);
        }

        // FIX 3: DUPLICATE CHECK ACROSS ALL ACTIVE BATCHES (VBA Logic)
        const existingMap = {};
        this.batches.forEach(b => {
            b.records.forEach(r => {
                existingMap[r.invoiceNo] = b.batchName;
            });
        });

        const newRecords = [];
        const duplicateRecords = [];

        parsedRecords.forEach(rec => {
            if (existingMap[rec.invoiceNo]) {
                duplicateRecords.push({ record: rec, foundIn: existingMap[rec.invoiceNo] });
            } else {
                newRecords.push(rec);
            }
        });

        if (duplicateRecords.length > 0) {
            const sampleDupes = duplicateRecords.slice(0, 5).map(d => `${d.record.invoiceNo} (${d.foundIn})`).join("\n");
            const userChoice = prompt(
                `⚠️ DUPLICATE DO DETECTED!\n\n` +
                `New DO Available: ${newRecords.length}\n` +
                `Duplicate DO Found: ${duplicateRecords.length}\n\n` +
                `Sample Duplicates:\n${sampleDupes}\n\n` +
                `Choose Action:\n` +
                `1 = Create New Batch with ONLY new DOs\n` +
                `2 = Append new DOs to an existing Batch\n` +
                `3 = Cancel Import`, "1"
            );

            if (userChoice === "1" || userChoice === null) {
                // Continue creating new batch with newRecords
            } else if (userChoice === "2") {
                if (this.batches.length === 0) {
                    alert("[WARNING] No existing batch found to append to. Creating a new batch instead.");
                } else {
                    const batchListStr = this.batches.map((b, i) => `${i + 1}. ${b.batchName}`).join("\n");
                    const targetInput = prompt(`Choose destination batch:\n${batchListStr}\n\nEnter number:`, "1");
                    const targetIdx = parseInt(targetInput) - 1;

                    if (!isNaN(targetIdx) && this.batches[targetIdx]) {
                        this.batches[targetIdx].records.push(...(newRecords.length > 0 ? newRecords : parsedRecords));
                        this.saveToStorage();
                        this.renderUI();
                        alert(`[SUCCESS] Appended ${newRecords.length > 0 ? newRecords.length : parsedRecords.length} DO(s) to ${this.batches[targetIdx].batchName}!`);
                        return; // Done
                    } else {
                        alert("[CANCELLED] Invalid batch selected. Import cancelled.");
                        return;
                    }
                }
            } else {
                alert("[CANCELLED] Import cancelled by user.");
                return;
            }
        }

        // Determine Batch Name (Batch 01, Batch 02, etc.)
        const nextBatchNum = this.batches.length + 1;
        const batchName = `Batch ${String(nextBatchNum).padStart(2, '0')}`;

        // First Run: Prompt user to specify DO Summary file date (Matching BatachSummary.bas macro)
        if (this.batches.length === 0) {
            const tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            const dd = String(tomorrow.getDate()).padStart(2, '0');
            const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
            const yyyy = tomorrow.getFullYear();
            const defaultDateStr = `${dd}${mm}${yyyy}`;

            const userDate = prompt("Enter date for DO Summary (e.g. 03082026):", defaultDateStr) || defaultDateStr;
            this.targetFileName = `DO Summary List ${userDate.trim()}.xlsx`;
            this.generatorDate = userDate.trim();
        }

        // Read or auto-suggest Wave Number
        const prevWave = this.batches.length > 0 ? this.batches[this.batches.length - 1].waveNumber : "01";
        const waveInput = prompt(`Enter Wave Number for ${batchName}:`, prevWave) || prevWave;

        this.batches.push({
            batchName: batchName,
            waveNumber: waveInput.trim(),
            records: newRecords.length > 0 ? newRecords : parsedRecords
        });

        this.currentBatchIndex = this.batches.length - 1;
        this.saveToStorage();
        this.renderUI();

        alert(`[SUCCESS] Created ${batchName} with ${newRecords.length > 0 ? newRecords.length : parsedRecords.length} DO records!`);
    }

    // Secondary file handler to bulk-update missing Route or Volume (m3) across existing batches
    async handleUpdateFileUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        if (this.batches.length === 0) {
            alert("[WARNING] No batches loaded to update! Please import your main SONY DO Summary List first.");
            event.target.value = "";
            return;
        }

        let updatedCount = 0;
        let routeUpdatedCount = 0;
        let volUpdatedCount = 0;

        for (const file of files) {
            const dataBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(dataBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            if (!rawRows || rawRows.length < 2) continue;

            const headerRow = rawRows[0].map(h => String(h || "").trim().toUpperCase());
            const colMap = {
                invoiceNo: this.findColIdx(headerRow, ["INVOICE NO.", "INVOICE", "INVOICE_NO", "DO NO", "DO_NUMBER", "SEQ"]),
                route:     this.findColIdx(headerRow, ["ROUTE", "ZONE"]),
                volume:    this.findColIdx(headerRow, ["VOLUME", "M3", "VOLUME (M3)"])
            };

            const invoiceIdx = colMap.invoiceNo !== -1 ? colMap.invoiceNo : 0;
            const routeIdx = colMap.route !== -1 ? colMap.route : 3;
            const volIdx = colMap.volume !== -1 ? colMap.volume : 8;

            for (let i = 1; i < rawRows.length; i++) {
                const row = rawRows[i];
                if (!row || row.length === 0) continue;

                const invoiceNo = String(row[invoiceIdx] || "").trim();
                if (!invoiceNo) continue;

                const rawRoute = String(row[routeIdx] || "").trim();
                const rawVolStr = String(row[volIdx] ?? "").trim();
                const cleanedVolStr = rawVolStr.toLowerCase().replace(/m3|m³/g, "").trim();
                const volVal = parseFloat(cleanedVolStr) || 0;

                // Match against all existing batches in memory
                this.batches.forEach(b => {
                    b.records.forEach(r => {
                        if (r.invoiceNo === invoiceNo) {
                            let isUpdated = false;

                            // Update Route if valid
                            if (rawRoute && rawRoute !== "-" && rawRoute.toUpperCase() !== "MISSING") {
                                if (r.route !== rawRoute || r.missingRoute) {
                                    r.route = rawRoute.toUpperCase();
                                    r.missingRoute = false;
                                    routeUpdatedCount++;
                                    isUpdated = true;
                                }
                            }

                            // Update Volume if > 0
                            if (volVal > 0) {
                                if (r.volume !== volVal || r.missingVolume) {
                                    r.volume = volVal;
                                    r.missingVolume = false;
                                    volUpdatedCount++;
                                    isUpdated = true;
                                }
                            }

                            if (isUpdated) updatedCount++;
                        }
                    });
                });
            }
        }

        event.target.value = ""; // Reset file picker

        if (updatedCount > 0) {
            this.saveToStorage();
            this.renderUI();
            alert(`[SUCCESS] Updated missing information for ${updatedCount} DO record(s):\n- ${routeUpdatedCount} Route(s) updated\n- ${volUpdatedCount} Volume(s) updated`);
        } else {
            alert("[INFO] No matching DO records were updated. Make sure INVOICE No. matches your loaded batches.");
        }
    }

    findColIdx(headers, candidates) {
        for (const candidate of candidates) {
            const idx = headers.indexOf(candidate);
            if (idx !== -1) return idx;
        }
        return -1;
    }

    // Compile Final Summary List by merging all batches
    async compileFinalSummary() {
        if (this.batches.length === 0) {
            alert("[INFO] No batches available to compile. Please upload a source file first!");
            return;
        }

        this.hasCompiledFinalSummary = true;

        const allRecords = [];
        this.batches.forEach(b => {
            b.records.filter(r => r.selected !== false).forEach(r => {
                allRecords.push({ ...r, batchOrigin: b.batchName, waveNumber: b.waveNumber });
            });
        });

        alert(`[SUCCESS] Compiled Final Summary List containing ${allRecords.length} total DOs across ${this.batches.length} batch(es)!`);
        this.saveToStorage();
        this.renderUI();
        await this.exportToExcel();
        this.syncToActivityTrend(allRecords);
    }

    // Auto-sync the compiled Final Summary totals into the DO Activity Trend history
    syncToActivityTrend(allRecords) {
        try {
            if (!allRecords || allRecords.length === 0) return;

            let dateStr = "";
            const nameMatch = (this.targetFileName || "").match(/(\d{2})(\d{2})(\d{4})/);
            if (nameMatch) {
                dateStr = `${nameMatch[3]}-${nameMatch[2]}-${nameMatch[1]}`;
            } else {
                const savedDate = (this.generatorDate || "").trim().match(/(\d{2})(\d{2})(\d{4})/);
                if (savedDate) {
                    dateStr = `${savedDate[3]}-${savedDate[2]}-${savedDate[1]}`;
                } else {
                    const targetDate = new Date();
                    targetDate.setDate(targetDate.getDate() + 1);
                    const yyyy = targetDate.getFullYear();
                    const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
                    const dd = String(targetDate.getDate()).padStart(2, '0');
                    dateStr = `${yyyy}-${mm}-${dd}`;
                }
            }
            if (!dateStr) return;

            let doCount = 0, skuTotal = 0, volTotal = 0, qtyTotal = 0;
            let lclDo = 0, lclSku = 0, lclVol = 0, lclQty = 0;

            allRecords.forEach(r => {
                doCount++;
                skuTotal += r.sku || 0;
                volTotal += r.volume || 0;
                qtyTotal += r.qty || 0;
                if ((r.remark || "").toUpperCase().includes("LCL")) {
                    lclDo++;
                    lclSku += r.sku || 0;
                    lclVol += r.volume || 0;
                    lclQty += r.qty || 0;
                }
            });

            let history = [];
            try {
                history = JSON.parse(localStorage.getItem("DO_Activity_Trend_History")) || [];
            } catch (e) {
                history = [];
            }
            if (!Array.isArray(history)) history = [];

            const entry = {
                date: dateStr,
                do: doCount,
                sku: skuTotal,
                vol: volTotal,
                qty: qtyTotal,
                lclDo, lclSku, lclVol, lclQty
            };

            const existingIdx = history.findIndex(h => h.date === dateStr);
            if (existingIdx !== -1) {
                history[existingIdx] = entry;
            } else {
                history.push(entry);
            }
            history.sort((a, b) => a.date.localeCompare(b.date));

            localStorage.setItem("DO_Activity_Trend_History", JSON.stringify(history));
            alert(`[SYNC] DO Activity Trend updated for ${dateStr} (${doCount} DO, ${skuTotal} SKU, ${volTotal.toFixed(2)} m³, ${qtyTotal} QTY).`);
        } catch (err) {
            console.error("Activity Trend sync failed:", err);
        }
    }

    // Export complete Multi-Sheet Workbook
    async exportToExcel() {
        if (this.batches.length === 0) {
            alert("[WARNING] No batch data to export! Please upload source files first.");
            return;
        }

        const workbook = XLSX.utils.book_new();

        const styleHeader = {
            font: { name: "Aptos Narrow", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "E31B23" } },
            alignment: { horizontal: "center", vertical: "center" }
        };

        const styleBatchSummaryHeader = {
            font: { name: "Aptos Narrow", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "E31B23" } },
            alignment: { horizontal: "center", vertical: "center" }
        };

        const styleBoldData = { font: { name: "Aptos Narrow", sz: 11, bold: true } };
        const styleNormalData = { font: { name: "Aptos Narrow", sz: 11, bold: false } };

        const styleRemarkHighlight = {
            font: { name: "Aptos Narrow", sz: 11, bold: true, color: { rgb: "9C0006" } },
            fill: { fgColor: { rgb: "FFC7CE" } },
            alignment: { horizontal: "left", vertical: "center" }
        };

        // 1. Add individual Batch sheets
        this.batches.forEach((b, batchIdx) => {
            const sheetData = [];
            const activeRecords = b.records.filter(r => r.selected !== false);
            
            const totalVol = activeRecords.reduce((sum, r) => sum + r.volume, 0);
            const totalQty = activeRecords.reduce((sum, r) => sum + r.qty, 0);
            const totalSku = activeRecords.reduce((sum, r) => sum + r.sku, 0);

            sheetData.push([
                `${activeRecords.length} DO`, "", "", "", `${b.batchName.toUpperCase()} SUMMARY`, "", "", "",
                parseFloat(totalVol.toFixed(3)), totalQty, totalSku, `Wave : ${b.waveNumber}`
            ]);

            sheetData.push([
                "INVOICE No.", "DIVISION", "SHP_CODE", "ROUTE",
                "CONSIGNEE_NAME", "ADDRESS1", "ADDRESS2", "ADDRESS3",
                "VOLUME", "QTY", "TOTAL SKU", "REMARK"
            ]);

            activeRecords.forEach(r => {
                sheetData.push([
                    r.invoiceNo, r.division, r.shpCode, r.route,
                    r.consignee, r.addr1, r.addr2, r.addr3,
                    r.volume, r.qty, r.sku, r.remark
                ]);
            });

            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            const lastR = activeRecords.length + 2;

            ws['A1'] = { f: `SUBTOTAL(103,A3:A${lastR})`, z: '0 "DO"' };
            ws['E1'] = { v: `${b.batchName.toUpperCase()} SUMMARY` };
            ws['I1'] = { f: `SUBTOTAL(109,I3:I${lastR})`, z: '0.00000 "M3"' };
            ws['J1'] = { f: `SUBTOTAL(109,J3:J${lastR})`, z: '0 "QTY"' };
            ws['K1'] = { f: `SUBTOTAL(109,K3:K${lastR})`, z: '0 "SKU"' };
            ws['L1'] = { v: `Wave : ${b.waveNumber}` };

            ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'].forEach(col => {
                const cellRef1 = `${col}1`;
                if (!ws[cellRef1]) ws[cellRef1] = { v: "" };
                ws[cellRef1].s = styleBatchSummaryHeader;

                const cellRef2 = `${col}2`;
                if (ws[cellRef2]) ws[cellRef2].s = styleHeader;
            });

            activeRecords.forEach((rec, rIdx) => {
                const rowNum = rIdx + 3;
                ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'].forEach(col => {
                    const cellRef = `${col}${rowNum}`;
                    if (ws[cellRef]) {
                        if (col === 'A') {
                            ws[cellRef].s = styleBoldData;
                        } else if (col === 'D' && (rec.missingRoute || !rec.route || rec.route.trim() === "" || rec.route === "-" || rec.route.toUpperCase() === "MISSING")) {
                            ws[cellRef].s = styleRemarkHighlight;
                        } else if (col === 'I' && (rec.missingVolume || rec.volume <= 0)) {
                            ws[cellRef].s = styleRemarkHighlight;
                        } else if (col === 'L' && rec.remark && rec.remark.trim() !== "" && rec.remark !== "-") {
                            ws[cellRef].s = styleRemarkHighlight;
                        } else {
                            ws[cellRef].s = styleNormalData;
                        }
                    }
                });
            });

            ws['!cols'] = [
                { wch: 16.22 },                 // A: INVOICE No.
                { wch: 5.66 },                  // B: DIVISION
                { wch: 14.55 },                 // C: SHP_CODE
                { wch: 6.44 },                  // D: ROUTE
                { wch: 30.78 },                 // E: CONSIGNEE_NAME
                { wch: 30.78 },                 // F: ADDRESS1
                { wch: 30.78, hidden: true },   // G: ADDRESS2
                { wch: 22.44, hidden: true },   // H: ADDRESS3
                { wch: 12.44 },                 // I: VOLUME
                { wch: 11.33 },                 // J: QTY
                { wch: 10.33 },                 // K: TOTAL SKU
                { wch: 12.33 }                  // L: REMARK
            ];

            ws['!autofilter'] = { ref: `A2:L${lastR}` };
            ws['!freeze'] = { xSplit: 0, ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' };
            ws['!tabColor'] = { rgb: (batchIdx % 2 === 0) ? "E31B23" : "FFC7CE" };

            XLSX.utils.book_append_sheet(workbook, ws, b.batchName);
        });

        // 2. Add Final Summary List Sheet if compiled
        if (this.hasCompiledFinalSummary) {
            const allFinalRecords = [];
            const waveDict = {};
            let challengerDOCount = 0;

            this.batches.forEach(b => {
                const batchCleanName = b.batchName.replace("Batch ", "");
                const waveCleanNum = b.waveNumber || "-";

                if (waveDict[waveCleanNum]) {
                    waveDict[waveCleanNum] += `,${batchCleanName}`;
                } else {
                    waveDict[waveCleanNum] = batchCleanName;
                }

                b.records.filter(r => r.selected !== false).forEach(r => {
                    allFinalRecords.push({ ...r, batchOrigin: b.batchName, waveNumber: b.waveNumber });
                    if (r.consignee && r.consignee.toUpperCase().includes("CHALLENGER")) {
                        challengerDOCount++;
                    }
                });
            });
            const finalData = [];

            finalData.push([
                "", "", "", "", "FINAL SUMMARY LIST", "", "", "",
                0, 0, 0, "ALL WAVES"
            ]);

            finalData.push([
                "INVOICE No.", "DIVISION", "SHP_CODE", "ROUTE",
                "CONSIGNEE_NAME", "ADDRESS1", "ADDRESS2", "ADDRESS3",
                "VOLUME", "QTY", "TOTAL SKU", "REMARK"
            ]);

            allFinalRecords.forEach(r => {
                finalData.push([
                    r.invoiceNo, r.division, r.shpCode, r.route,
                    r.consignee, r.addr1, r.addr2, r.addr3,
                    r.volume, r.qty, r.sku, r.remark
                ]);
            });

            const wsFinal = XLSX.utils.aoa_to_sheet(finalData);
            const finalLastR = allFinalRecords.length + 2;

            wsFinal['A1'] = { f: `SUBTOTAL(103,A3:A${finalLastR})`, z: '0 "DO"' };
            wsFinal['E1'] = { v: "FINAL SUMMARY LIST" };
            wsFinal['I1'] = { f: `SUBTOTAL(109,I3:I${finalLastR})`, z: '0.00000 "M3"' };
            wsFinal['J1'] = { f: `SUBTOTAL(109,J3:J${finalLastR})`, z: '0 "QTY"' };
            wsFinal['K1'] = { f: `SUBTOTAL(109,K3:K${finalLastR})`, z: '0 "SKU"' };
            wsFinal['L1'] = { v: "" };

            ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'].forEach(col => {
                const cellRef1 = `${col}1`;
                if (!wsFinal[cellRef1]) wsFinal[cellRef1] = { v: "" };
                wsFinal[cellRef1].s = styleBatchSummaryHeader;

                const cellRef2 = `${col}2`;
                if (wsFinal[cellRef2]) wsFinal[cellRef2].s = styleHeader;
            });

            allFinalRecords.forEach((rec, rIdx) => {
                const rowNum = rIdx + 3;
                ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'].forEach(col => {
                    const cellRef = `${col}${rowNum}`;
                    if (wsFinal[cellRef]) {
                        if (col === 'A') {
                            wsFinal[cellRef].s = styleBoldData;
                        } else if (col === 'D' && (rec.missingRoute || !rec.route || rec.route.trim() === "" || rec.route === "-" || rec.route.toUpperCase() === "MISSING")) {
                            wsFinal[cellRef].s = styleRemarkHighlight;
                        } else if (col === 'I' && (rec.missingVolume || rec.volume <= 0)) {
                            wsFinal[cellRef].s = styleRemarkHighlight;
                        } else if (col === 'L' && rec.remark && rec.remark.trim() !== "" && rec.remark !== "-") {
                            wsFinal[cellRef].s = styleRemarkHighlight;
                        } else {
                            wsFinal[cellRef].s = styleNormalData;
                        }
                    }
                });
            });

            wsFinal['N3'] = { v: "WAVE NUMBER", s: {
                font: { name: "Aptos Narrow", sz: 13, bold: true, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "E31B23" } },
                alignment: { horizontal: "center", vertical: "center" }
            }};

            let dashRow = 4;
            Object.keys(waveDict).forEach(wave => {
                const cellRef = `N${dashRow}`;
                wsFinal[cellRef] = {
                    v: `  • Wave ${wave}  -->  Batch ${waveDict[wave]}`,
                    s: {
                        font: { name: "Aptos Narrow", sz: 11, bold: true, color: { rgb: "000000" } },
                        fill: { fgColor: { rgb: "FFEBEE" } },
                        alignment: { horizontal: "left", vertical: "center" }
                    }
                };
                dashRow++;
            });

            const challengerCellRef = `N${dashRow + 1}`;
            wsFinal[challengerCellRef] = {
                v: `CHALLENGER DO : ${challengerDOCount}`,
                s: {
                    font: { name: "Aptos Narrow", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
                    fill: { fgColor: { rgb: "E31B23" } },
                    alignment: { horizontal: "center", vertical: "center" }
                }
            };

            wsFinal['!cols'] = [
                { wch: 16.22 }, { wch: 5.66 }, { wch: 14.55 }, { wch: 6.44 },
                { wch: 30.78 }, { wch: 30.78 }, { wch: 30.78, hidden: true }, { wch: 22.44, hidden: true },
                { wch: 12.44 }, { wch: 11.33 }, { wch: 10.33 }, { wch: 12.33 },
                { wch: 4.00 }, { wch: 32.00 }
            ];

            const maxRow = Math.max(finalLastR, dashRow + 1);
            wsFinal['!ref'] = `A1:N${maxRow}`;
            wsFinal['!autofilter'] = { ref: `A2:L${finalLastR}` };
            wsFinal['!freeze'] = { xSplit: 0, ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' };
            wsFinal['!tabColor'] = { rgb: "FF0000" };

            XLSX.utils.book_append_sheet(workbook, wsFinal, "Final Summary List");
        }

        let exportFileName = this.targetFileName || "";
        if (!exportFileName) {
            const tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            const dd = String(tomorrow.getDate()).padStart(2, '0');
            const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
            const yyyy = tomorrow.getFullYear();
            exportFileName = `DO Summary List ${dd}${mm}${yyyy}.xlsx`;
        }
        if (!exportFileName.toLowerCase().endsWith(".xlsx")) {
            exportFileName += ".xlsx";
        }

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

    renderUI() {
        this.renderKPIs();
        this.renderBatchTabs();
        this.renderEmailTargets();
        this.renderTable();
    }

    renderEmailTargets() {
        const select = document.getElementById("emailTargetSelect");
        if (!select) return;

        let options = "";
        if (this.hasCompiledFinalSummary) {
            options += `<option value="FINAL">Final Summary</option>`;
        }
        this.batches.forEach((b, idx) => {
            options += `<option value="batch-${idx}">${b.batchName}</option>`;
        });

        if (options === "") {
            options = `<option value="">No batch</option>`;
            select.innerHTML = options;
            return;
        }

        const prev = select.value;
        select.innerHTML = options;

        if (prev && Array.from(select.options).some(o => o.value === prev)) {
            select.value = prev;
        } else {
            select.value = this.hasCompiledFinalSummary && this.batches.length === 0
                ? "FINAL"
                : `batch-${this.currentBatchIndex}`;
        }
    }

    onEmailTargetChange(value) {
        if (value && value.startsWith("batch-")) {
            const idx = parseInt(value.replace("batch-", ""), 10);
            if (!isNaN(idx) && this.batches[idx]) {
                this.currentBatchIndex = idx;
                this.renderTable();
            }
        }
    }

    async generateEmail() {
        const select = document.getElementById("emailTargetSelect");
        const target = select ? select.value : "";
        if (!target) {
            alert("[WARNING] No batch available to generate email. Please import a source file first.");
            return;
        }

        let batchLabel = "";
        let doCount = 0;
        let skuTotal = 0;
        let doQty = 0;

        if (target === "FINAL") {
            batchLabel = "FINAL";
            this.batches.forEach(b => {
                b.records.filter(r => r.selected !== false).forEach(r => {
                    doCount++;
                    skuTotal += r.sku;
                    doQty += r.qty;
                });
            });
        } else {
            const idx = parseInt(target.replace("batch-", ""), 10);
            const batch = this.batches[idx];
            if (!batch) {
                alert("[ERROR] Selected batch could not be found.");
                return;
            }
            batchLabel = batch.batchName;
            batch.records.filter(r => r.selected !== false).forEach(r => {
                doCount++;
                skuTotal += r.sku;
                doQty += r.qty;
            });
        }

        const emailText = "Dear all,\r\n\r\n" +
            "Kindly refer attached file for D/O Summary.\r\n\r\n" +
            `Total DO for ${batchLabel} = ${doCount} (total SKU & QTY: ${skuTotal}/${doQty})\r\n\r\n` +
            "Please check if any DO, Qty & Remarks not tally with your data.";

        try {
            await navigator.clipboard.writeText(emailText);
        } catch (e) {
            const ta = document.createElement("textarea");
            ta.value = emailText;
            ta.style.position = "fixed";
            ta.style.opacity = "0";
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
        }

        alert("Email text copied to clipboard:\n\n" + emailText);
    }

    renderKPIs() {
        let totalBatches = this.batches.length;
        let totalDO = 0;
        let totalVol = 0;
        let totalQty = 0;
        let totalSku = 0;

        this.batches.forEach(b => {
            const activeRecs = b.records.filter(r => r.selected === true);
            totalDO += activeRecs.length;
            activeRecs.forEach(r => {
                totalVol += r.volume;
                totalQty += r.qty;
                totalSku += r.sku;
            });
        });

        const elBatches = document.getElementById("gen-kpi-batches");
        const elDO = document.getElementById("gen-kpi-do");
        const elVol = document.getElementById("gen-kpi-vol");
        const elQty = document.getElementById("gen-kpi-qty");
        const elSku = document.getElementById("gen-kpi-sku");

        if (elBatches) elBatches.innerText = totalBatches;
        if (elDO) elDO.innerText = totalDO;
        if (elVol) elVol.innerText = `${totalVol.toFixed(2)} m³`;
        if (elQty) elQty.innerText = totalQty.toLocaleString();
        if (elSku) elSku.innerText = totalSku.toLocaleString();
    }

    renderBatchTabs() {
        const container = document.getElementById("generatorBatchTabs");
        if (!container) return;

        if (this.batches.length === 0) {
            container.innerHTML = '<span style="color:#71717a; font-size:13px;">No batches created yet. Upload a source file to start.</span>';
            return;
        }

        let html = "";
        this.batches.forEach((b, idx) => {
            const isActive = idx === this.currentBatchIndex;
            const activeClass = isActive ? "active-tab" : "";
            html += `<button class="batch-tab-btn ${activeClass}" onclick="summaryGenerator.selectBatch(${idx})">${b.batchName} (${b.records.length} DO)</button>`;
        });

        container.innerHTML = html;
    }

    selectBatch(index) {
        this.currentBatchIndex = index;
        this.renderUI();
    }

    toggleSelectAll(checked) {
        if (this.batches.length === 0 || !this.batches[this.currentBatchIndex]) return;
        const activeBatch = this.batches[this.currentBatchIndex];
        const isSelected = Boolean(checked);

        activeBatch.records.forEach(r => {
            r.selected = isSelected;
        });

        this.saveToStorage();
        this.renderUI();
    }

    toggleRowSelect(recordIdx, checked) {
        if (this.batches[this.currentBatchIndex] && this.batches[this.currentBatchIndex].records[recordIdx]) {
            this.batches[this.currentBatchIndex].records[recordIdx].selected = Boolean(checked);
            this.saveToStorage();

            const activeBatch = this.batches[this.currentBatchIndex];
            const allChecked = activeBatch.records.length > 0 && activeBatch.records.every(r => r.selected === true);
            const mainCheckbox = document.getElementById("selectAllCheckbox");
            if (mainCheckbox) mainCheckbox.checked = allChecked;

            this.renderKPIs();
        }
    }

    addManualDO() {
        if (this.batches.length === 0) {
            const waveInput = prompt("Enter Wave Number for new Batch 01:", "01");
            if (waveInput === null) return;

            this.batches.push({
                batchName: "Batch 01",
                waveNumber: (waveInput || "01").trim(),
                records: []
            });
            this.currentBatchIndex = 0;

            const tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            const dd = String(tomorrow.getDate()).padStart(2, '0');
            const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
            const yyyy = tomorrow.getFullYear();
            const defaultDateStr = `${dd}${mm}${yyyy}`;
            const userDate = prompt("Enter date for DO Summary (e.g. 03082026):", defaultDateStr) || defaultDateStr;
            this.targetFileName = `DO Summary List ${userDate.trim()}.xlsx`;
            this.generatorDate = userDate.trim();
        }

        const activeBatch = this.batches[this.currentBatchIndex];
        const existingManualCount = activeBatch.records.filter(r => r.invoiceNo.startsWith("MANUAL-")).length;
        const nextNum = String(existingManualCount + 1).padStart(3, '0');

        const newRecord = {
            invoiceNo: `MANUAL-${nextNum}`,
            division: "",
            shpCode: "",
            route: "",
            consignee: "",
            addr1: "",
            addr2: "",
            addr3: "",
            volume: 0,
            qty: 0,
            sku: 0,
            remark: "",
            missingRoute: true,
            missingVolume: true,
            selected: false
        };

        activeBatch.records.push(newRecord);
        this.saveToStorage();
        this.renderUI();

        requestAnimationFrame(() => {
            const wrapper = document.querySelector('.compact-wrapper');
            if (wrapper) wrapper.scrollTop = wrapper.scrollHeight;
        });
    }

    deleteSelectedRows() {
        if (this.batches.length === 0 || !this.batches[this.currentBatchIndex]) return;
        const activeBatch = this.batches[this.currentBatchIndex];
        const selectedCount = activeBatch.records.filter(r => r.selected === true).length;
        
        if (selectedCount === 0) {
            alert("[INFO] No rows selected to delete.");
            return;
        }

        if (!confirm(`Are you sure you want to delete ${selectedCount} selected DO order(s) from ${activeBatch.batchName}?`)) return;

        activeBatch.records = activeBatch.records.filter(r => r.selected !== true);
        this.saveToStorage();
        this.renderUI();
        alert(`[SUCCESS] Deleted ${selectedCount} selected DO order(s).`);
    }

    updateRemark(batchIdx, recordIdx, val) {
        if (this.batches[batchIdx] && this.batches[batchIdx].records[recordIdx]) {
            this.batches[batchIdx].records[recordIdx].remark = val.trim();
            this.saveToStorage();
        }
    }

    updateRoute(batchIdx, recordIdx, val) {
        if (this.batches[batchIdx] && this.batches[batchIdx].records[recordIdx]) {
            const cleanVal = val.trim().toUpperCase();
            const rec = this.batches[batchIdx].records[recordIdx];
            rec.route = cleanVal;
            rec.missingRoute = !cleanVal || cleanVal === "-" || cleanVal === "MISSING";
            this.saveToStorage();
            this.renderKPIs();
            this.renderTable();
        }
    }

    updateVolume(batchIdx, recordIdx, val) {
        if (this.batches[batchIdx] && this.batches[batchIdx].records[recordIdx]) {
            const num = parseFloat(val) || 0;
            const rec = this.batches[batchIdx].records[recordIdx];
            rec.volume = num;
            rec.missingVolume = num <= 0;
            this.saveToStorage();
            this.renderKPIs();
            this.renderTable();
        }
    }

    setSearchQuery(val) {
        this.searchQuery = String(val || "").trim().toLowerCase();
        this.renderTable();
    }

    updateField(batchIdx, recordIdx, field, val) {
        if (this.batches[batchIdx] && this.batches[batchIdx].records[recordIdx]) {
            const rec = this.batches[batchIdx].records[recordIdx];
            if (field === 'qty' || field === 'sku') {
                rec[field] = parseInt(val) || 0;
            } else {
                rec[field] = val.trim();
            }
            this.saveToStorage();
            if (field === 'qty' || field === 'sku') {
                this.renderKPIs();
            }
        }
    }

    renderTable() {
        const tbody = document.getElementById("generatorTableBody");
        if (!tbody) return;

        const mainCheckbox = document.getElementById("selectAllCheckbox");
        const elSearchCount = document.getElementById("batchSearchCount");

        if (this.batches.length === 0 || !this.batches[this.currentBatchIndex]) {
            tbody.innerHTML = '<tr><td colspan="13" style="text-align:center; padding:24px; color:#71717a;">No batch data loaded. Please upload a source file.</td></tr>';
            if (mainCheckbox) mainCheckbox.checked = false;
            if (elSearchCount) elSearchCount.style.display = "none";
            return;
        }

        const activeBatch = this.batches[this.currentBatchIndex];
        const records = activeBatch.records;

        const filteredRecords = [];
        records.forEach((r, originalIdx) => {
            if (this.searchQuery) {
                const searchStr = `${r.invoiceNo} ${r.division} ${r.shpCode} ${r.route} ${r.consignee} ${r.addr1} ${r.addr2} ${r.addr3} ${r.remark}`.toLowerCase();
                if (!searchStr.includes(this.searchQuery)) return;
            }
            filteredRecords.push({ record: r, idx: originalIdx });
        });

        if (elSearchCount) {
            elSearchCount.style.display = "inline-block";
            if (this.searchQuery) {
                elSearchCount.innerText = `${filteredRecords.length} / ${records.length} DOs`;
            } else {
                elSearchCount.innerText = `${records.length} DOs`;
            }
        }

        const allChecked = records.length > 0 && records.every(r => r.selected === true);
        const noneChecked = records.length > 0 && records.every(r => r.selected === false);
        
        if (mainCheckbox) mainCheckbox.checked = allChecked;
        
        const btnSelectAll = document.getElementById("btnSelectAll");
        const btnDeselectAll = document.getElementById("btnDeselectAll");
        if (btnSelectAll && btnDeselectAll) {
            const setBtnState = (btn, checked, label) => {
                const icon = btn.querySelector('svg');
                const text = btn.querySelector('span');
                if (icon) {
                    if (checked) {
                        icon.innerHTML = '<polyline points="20 6 9 17 4 12"/>';
                    } else {
                        icon.innerHTML = '<rect x="3" y="3" width="18" height="18" rx="2"/>';
                    }
                }
                if (text) text.textContent = label;
            };

            if (allChecked) {
                setBtnState(btnSelectAll, true, "Select All");
                setBtnState(btnDeselectAll, false, "Deselect All");
            } else if (noneChecked) {
                setBtnState(btnSelectAll, false, "Select All");
                setBtnState(btnDeselectAll, true, "Deselect All");
            } else {
                setBtnState(btnSelectAll, false, "Select All");
                setBtnState(btnDeselectAll, false, "Deselect All");
            }
        }

        if (filteredRecords.length === 0) {
            tbody.innerHTML = `<tr><td colspan="13" style="text-align:center; padding:24px; color:#71717a;">No DO records match search "${this.searchQuery}".</td></tr>`;
            return;
        }

        const bi = this.currentBatchIndex;
        let html = "";
        filteredRecords.forEach((item) => {
            const r = item.record;
            const idx = item.idx;
            const isManual = r.invoiceNo.startsWith("MANUAL-");

            const isMissingRoute = r.missingRoute || !r.route || r.route.trim() === "" || r.route === "-" || r.route.toUpperCase() === "MISSING";
            const isMissingVol = r.missingVolume || r.volume <= 0;
            const isChecked = r.selected === true ? "checked" : "";

            const routeInputHtml = `<input type="text" class="compact-input ${isMissingRoute ? 'missing-highlight' : ''}" 
                value="${isMissingRoute ? '' : (r.route || '')}" placeholder="ROUTE" 
                onchange="summaryGenerator.updateRoute(${bi}, ${idx}, this.value)" 
                style="width: 75px; text-transform: uppercase;">`;

            const volInputHtml = `<input type="number" step="0.001" min="0" class="compact-input ${isMissingVol ? 'missing-highlight' : ''}" 
                value="${r.volume > 0 ? r.volume : ''}" placeholder="0.000" 
                onchange="summaryGenerator.updateVolume(${bi}, ${idx}, this.value)" 
                style="width: 80px; text-align: right;">`;

            if (isManual) {
                html += `<tr style="background: rgba(59, 130, 246, 0.05);">
                    <td style="text-align: center;">
                        <input type="checkbox" class="row-select-checkbox" ${isChecked} onchange="summaryGenerator.toggleRowSelect(${idx}, this.checked)" style="cursor: pointer;">
                    </td>
                    <td><input type="text" class="compact-input" value="${r.invoiceNo}" placeholder="DO Number" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'invoiceNo', this.value)" style="width: 110px; font-family: monospace; font-weight: 700; color: #60a5fa;"></td>
                    <td><input type="text" class="compact-input" value="${r.division || ''}" placeholder="DIV" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'division', this.value)" style="width: 50px;"></td>
                    <td><input type="text" class="compact-input" value="${r.shpCode || ''}" placeholder="SHP" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'shpCode', this.value)" style="width: 70px;"></td>
                    <td>${routeInputHtml}</td>
                    <td><input type="text" class="compact-input" value="${r.consignee || ''}" placeholder="Consignee" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'consignee', this.value)" style="width: 140px; font-weight: 600;"></td>
                    <td><input type="text" class="compact-input" value="${r.addr1 || ''}" placeholder="Address 1" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'addr1', this.value)" style="width: 120px;"></td>
                    <td><input type="text" class="compact-input" value="${r.addr2 || ''}" placeholder="Address 2" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'addr2', this.value)" style="width: 120px;"></td>
                    <td><input type="text" class="compact-input" value="${r.addr3 || ''}" placeholder="Address 3" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'addr3', this.value)" style="width: 100px;"></td>
                    <td>${volInputHtml} m³</td>
                    <td><input type="number" class="compact-input" value="${r.qty || ''}" placeholder="0" min="0" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'qty', this.value)" style="width: 60px; text-align: right;"></td>
                    <td><input type="number" class="compact-input" value="${r.sku || ''}" placeholder="0" min="0" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'sku', this.value)" style="width: 50px; text-align: right;"></td>
                    <td>
                        <input type="text" class="compact-remark-input" value="${r.remark || ''}" placeholder="Enter remark..." 
                               onchange="summaryGenerator.updateRemark(${bi}, ${idx}, this.value)">
                    </td>
                </tr>`;
            } else {
                html += `<tr>
                    <td style="text-align: center;">
                        <input type="checkbox" class="row-select-checkbox" ${isChecked} onchange="summaryGenerator.toggleRowSelect(${idx}, this.checked)" style="cursor: pointer;">
                    </td>
                    <td><input type="text" class="compact-input" value="${r.invoiceNo}" placeholder="DO Number" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'invoiceNo', this.value)" style="width: 110px; font-family: monospace; font-weight: 700; color: #60a5fa;"></td>
                    <td><input type="text" class="compact-input" value="${r.division || ''}" placeholder="DIV" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'division', this.value)" style="width: 50px;"></td>
                    <td><input type="text" class="compact-input" value="${r.shpCode || ''}" placeholder="SHP" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'shpCode', this.value)" style="width: 70px;"></td>
                    <td>${routeInputHtml}</td>
                    <td><input type="text" class="compact-input" value="${r.consignee || ''}" placeholder="Consignee" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'consignee', this.value)" style="width: 140px; font-weight: 600;"></td>
                    <td><input type="text" class="compact-input" value="${r.addr1 || ''}" placeholder="Address 1" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'addr1', this.value)" style="width: 120px;"></td>
                    <td><input type="text" class="compact-input" value="${r.addr2 || ''}" placeholder="Address 2" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'addr2', this.value)" style="width: 120px;"></td>
                    <td><input type="text" class="compact-input" value="${r.addr3 || ''}" placeholder="Address 3" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'addr3', this.value)" style="width: 100px;"></td>
                    <td>${volInputHtml} m³</td>
                    <td><input type="number" class="compact-input" value="${r.qty || ''}" placeholder="0" min="0" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'qty', this.value)" style="width: 60px; text-align: right;"></td>
                    <td><input type="number" class="compact-input" value="${r.sku || ''}" placeholder="0" min="0" 
                        onchange="summaryGenerator.updateField(${bi}, ${idx}, 'sku', this.value)" style="width: 50px; text-align: right;"></td>
                    <td>
                        <input type="text" class="compact-remark-input" value="${r.remark || ''}" placeholder="Enter remark..." 
                               onchange="summaryGenerator.updateRemark(${bi}, ${idx}, this.value)">
                    </td>
                </tr>`;
            }
        });

        tbody.innerHTML = html;
    }
}

// Global Instance
let summaryGenerator = null;
document.addEventListener("DOMContentLoaded", () => {
    summaryGenerator = new DOSummaryGenerator();
});