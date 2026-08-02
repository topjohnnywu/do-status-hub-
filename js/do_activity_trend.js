// DO Activity Trend - JS Conversion of BulkImportDailyReports.bas

class DOActivityTrend {
    constructor() {
        this.history = []; // [{ date: 'YYYY-MM-DD', do, sku, vol, qty, lclDo, lclSku, lclVol, lclQty }]
        this.timeframe = "30";
        this.chartType = "main";
        this.searchQuery = "";
        this.chart = null;
        this.currentPage = 1;
        this.pageSize = 10; // Default 10 per page
        this.init();
    }

    init() {
        this.loadFromStorage();
        this.renderUI();
    }

    loadFromStorage() {
        const stored = localStorage.getItem("DO_Activity_Trend_History");
        if (stored) {
            try {
                this.history = JSON.parse(stored);
                if (!Array.isArray(this.history)) this.history = [];
            } catch (e) {
                console.error("Failed to load DO Activity Trend data", e);
                this.history = [];
            }
        }
        const tf = localStorage.getItem("DO_Activity_Trend_Timeframe");
        if (tf) this.timeframe = tf;
        const select = document.getElementById("datTimeframeFilter");
        if (select) select.value = this.timeframe;

        const ct = localStorage.getItem("DO_Activity_Trend_ChartType");
        if (ct) this.chartType = ct;
        const chartTypeSelect = document.getElementById("datChartTypeFilter");
        if (chartTypeSelect) chartTypeSelect.value = this.chartType;
    }

    saveToStorage() {
        localStorage.setItem("DO_Activity_Trend_History", JSON.stringify(this.history));
        localStorage.setItem("DO_Activity_Trend_Timeframe", this.timeframe);
        localStorage.setItem("DO_Activity_Trend_ChartType", this.chartType);
    }

    reset() {
        if (this.history.length > 0 && !confirm("Are you sure you want to reset all activity history?")) return;
        this.history = [];
        localStorage.removeItem("DO_Activity_Trend_History");
        localStorage.removeItem("DO_Activity_Trend_Timeframe");
        localStorage.removeItem("DO_Activity_Trend_ChartType");
        const fileInput = document.getElementById("activityFilePicker");
        if (fileInput) fileInput.value = "";
        this.renderUI();
    }

    // Main entry point for uploading daily operations reports (Clean Slate mode)
    async handleFileUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        // Clean Slate: wipe old data before import (matching BulkImportDailyReports.bas)
        if (this.history.length > 0) {
            const wipe = confirm("Clean Slate enabled: existing history will be wiped before importing the new files.\n\nContinue?");
            if (!wipe) {
                event.target.value = "";
                return;
            }
            this.history = [];
        }

        const summary = await this.processFiles(files);
        event.target.value = ""; // Reset file picker
        this.alertSummary(summary);
    }

    // Update-only mode: upserts by date without wiping existing history
    async handleUpdateFileUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        const summary = await this.processFiles(files);
        event.target.value = ""; // Reset file picker
        this.alertSummary(summary, true);
    }

    async processFiles(files) {
        let successCount = 0;
        let failureCount = 0;
        let failedFilesList = "";

        for (const file of files) {
            const result = await this.processSourceFile(file);
            if (result.ok) {
                successCount++;
            } else {
                failureCount++;
                failedFilesList += `\n   - ${file.name} (${result.reason})`;
            }
        }

        this.saveToStorage();
        this.renderUI();
        return { successCount, failureCount, failedFilesList };
    }

    alertSummary(summary, isUpdate = false) {
        let finalMessage = isUpdate
            ? `[UPDATE COMPLETED]\n-- Successfully Updated: ${summary.successCount} files.\n-- Failed/Skipped: ${summary.failureCount} files.`
            : `[PROCESS COMPLETED]\n-- Successfully Processed: ${summary.successCount} files.\n-- Failed/Skipped: ${summary.failureCount} files.`;
        if (summary.failureCount > 0) {
            finalMessage += `\n\nThe following files were skipped:${summary.failedFilesList}`;
        }
        alert(finalMessage);
    }

    async processSourceFile(file) {
        try {
            const fileName = file.name;

            // Date extraction from filename (matching ExtractDateFromFilename)
            let dateStr = this.extractDateFromFilename(fileName);
            if (!dateStr) {
                const manual = prompt(`No valid 8-digit date (ddmmyyyy) found in "${fileName}".\nEnter date (e.g. 03082026):`, "");
                if (!manual) return { ok: false, reason: "Date Error" };
                const parsed = this.parseDateString(manual.trim());
                if (!parsed) return { ok: false, reason: "Date Error" };
                dateStr = parsed;
            }

            const dataBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(dataBuffer, { type: 'array' });

            // Find "Final Summary List" sheet (matching VBA sheet lookup)
            let sheetName = workbook.SheetNames.find(name => {
                const lower = name.toLowerCase();
                return lower.includes("final summary") || lower.includes("final");
            }) || workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) return { ok: false, reason: "Missing Sheet" };

            const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
            if (rawRows.length < 1) return { ok: false, reason: "No Data" };

            // --- SMART 3-PASS LAYOUT PARSER ---
            // 1. Header Row Finder (detect if header is at Row 0 or Row 1)
            let headerIdx = -1;
            for (let i = 0; i < Math.min(5, rawRows.length); i++) {
                const rStr = rawRows[i].map(c => String(c ?? "").toUpperCase()).join(" ");
                if (rStr.includes("INVOICE") || rStr.includes("ROUTE")) {
                    headerIdx = i;
                    break;
                }
            }
            if (headerIdx === -1) headerIdx = 0;

            let doPackage = 0;
            let spaceOccupied = 0;
            let totalQty = 0;
            let skuAssortment = 0;
            let hasSummaryRow = false;
            let isBottomSummary = false;

            // Pass A: Top Summary Row Inspection (Row 0 if Header is Row 1, e.g. 13072026.xlsx)
            if (headerIdx > 0) {
                const topRow = rawRows[0] || [];
                const v8 = this.parseFirstNumber(topRow[8]);
                const v9 = this.parseFirstNumber(topRow[9]);
                const v0 = this.parseFirstNumber(topRow[0]);
                if ((v8 > 0 || v9 > 0) && v0 < 10000) {
                    doPackage = v0;
                    spaceOccupied = v8;
                    totalQty = v9;
                    skuAssortment = this.parseFirstNumber(topRow[10]);
                    hasSummaryRow = true;
                }
            }

            // Pass B: Bottom Summary Row Inspection (Last row, e.g. 01072026.xlsx)
            if (!hasSummaryRow && rawRows.length > headerIdx + 1) {
                const lastRow = rawRows[rawRows.length - 1] || [];
                const v8 = this.parseFirstNumber(lastRow[8]);
                const v9 = this.parseFirstNumber(lastRow[9]);
                const v0 = this.parseFirstNumber(lastRow[0]);
                if ((v8 > 0 || v9 > 0) && v0 < 10000) {
                    doPackage = v0;
                    spaceOccupied = v8;
                    totalQty = v9;
                    skuAssortment = this.parseFirstNumber(lastRow[10]);
                    hasSummaryRow = true;
                    isBottomSummary = true;
                }
            }

            // Pass C: Data Row Summation & LCL Transactional Scanner
            let calcDo = 0, calcVol = 0, calcQty = 0, calcSku = 0;
            let lclDo = 0, lclVol = 0, lclQty = 0, lclSku = 0;

            const dataEndIdx = (hasSummaryRow && isBottomSummary) ? rawRows.length - 1 : rawRows.length;

            for (let i = headerIdx + 1; i < dataEndIdx; i++) {
                const row = rawRows[i];
                if (!row || row.length === 0) continue;

                const inv = String(row[0] ?? row[1] ?? "").trim();
                if (!inv || inv.toUpperCase().includes("TOTAL") || inv.toUpperCase().includes("INVOICE")) continue;

                const v8 = this.parseFirstNumber(row[8]);  // Vol (Col I)
                const v9 = this.parseFirstNumber(row[9]);  // Qty (Col J)
                const v10 = this.parseFirstNumber(row[10]); // SKU (Col K)

                calcDo++;
                calcVol += v8;
                calcQty += v9;
                calcSku += v10;

                const rem = String(row[11] ?? "").toUpperCase(); // Remark (Col L)
                if (rem.includes("LCL")) {
                    lclDo++;
                    lclVol += v8;
                    lclQty += v9;
                    lclSku += v10;
                }
            }

            // Fallback to computed row sums if no summary row was found or metrics were invalid/zero
            if (!hasSummaryRow || doPackage <= 0 || doPackage > 10000 || spaceOccupied === 0) {
                doPackage = calcDo;
                spaceOccupied = calcVol;
                totalQty = calcQty;
                skuAssortment = calcSku;
            }

            // Upsert into history (matching dedup-by-date logic)
            const existingIdx = this.history.findIndex(h => h.date === dateStr);
            const entry = {
                date: dateStr,
                do: doPackage,
                sku: skuAssortment,
                vol: spaceOccupied,
                qty: totalQty,
                lclDo, lclSku, lclVol, lclQty
            };
            if (existingIdx !== -1) {
                this.history[existingIdx] = entry;
            } else {
                this.history.push(entry);
            }

            this.history.sort((a, b) => a.date.localeCompare(b.date));
            return { ok: true };
        } catch (err) {
            console.error("File processing error:", err);
            return { ok: false, reason: "Parse Error" };
        }
    }

    // Extract date from filename - finds 8 consecutive digits (ddmmyyyy)
    extractDateFromFilename(fileName) {
        const cleanName = fileName.replace(/\.[^.]+$/, "");
        const match = cleanName.match(/\d{8}/);
        if (!match) return null;
        return this.parseDateString(match[0]);
    }

    parseDateString(str) {
        const digits = String(str).replace(/\D/g, "").slice(0, 8);
        if (digits.length !== 8) return null;
        const day = parseInt(digits.slice(0, 2), 10);
        const month = parseInt(digits.slice(2, 4), 10);
        const year = parseInt(digits.slice(4, 8), 10);
        if (month < 1 || month > 12 || day < 1 || day > 31 || year < 2000 || year > 2099) return null;
        const d = new Date(year, month - 1, day);
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }

    // Parse first number from a mixed cell value (matching ParseFirstNumber)
    parseFirstNumber(val) {
        if (val === "" || val === null || val === undefined) return 0;
        if (typeof val === 'number') return val;
        const m = String(val).replace(/\u00A0/g, " ").match(/-?\d+(\.\d+)?/);
        return m ? parseFloat(m[0]) : 0;
    }

    setTimeframe(val) {
        this.timeframe = val;
        this.currentPage = 1;
        this.saveToStorage();
        this.renderChart();
        this.renderTable();
    }

    setChartType(val) {
        this.chartType = val;
        this.saveToStorage();
        this.renderChart();
    }

    setSearchQuery(val) {
        this.searchQuery = String(val || "").trim().toLowerCase();
        this.currentPage = 1;
        this.renderTable();
    }

    updateTimeframeDropdown() {
        const select = document.getElementById("datTimeframeFilter");
        if (!select) return;

        const monthMap = new Map();
        const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

        this.history.forEach(h => {
            if (!h.date) return;
            const [y, m] = h.date.split("-");
            if (y && m) {
                const key = `${y}-${m}`;
                const mIdx = parseInt(m, 10) - 1;
                if (mIdx >= 0 && mIdx < 12) {
                    monthMap.set(key, `${monthNames[mIdx]} ${y}`);
                }
            }
        });

        let html = `
            <option value="30">Last 30 Days</option>
            <option value="7">Last 7 Days</option>
            <option value="all">All Time</option>
        `;

        if (monthMap.size > 0) {
            html += `<optgroup label="Monthly History">`;
            const sortedKeys = Array.from(monthMap.keys()).sort();
            sortedKeys.forEach(k => {
                html += `<option value="month-${k}">${monthMap.get(k)}</option>`;
            });
            html += `</optgroup>`;
        }

        select.innerHTML = html;
        if (this.timeframe) select.value = this.timeframe;
    }

    getFilteredHistory() {
        if (this.history.length === 0) return [];
        if (this.timeframe === "all") return this.history.slice();
        if (this.timeframe === "7") return this.history.slice(-7);
        if (this.timeframe === "30") return this.history.slice(-30);

        if (this.timeframe.startsWith("month-")) {
            const ym = this.timeframe.replace("month-", "");
            return this.history.filter(h => h.date.startsWith(ym));
        }

        return this.history.slice(-30);
    }

    renderUI() {
        this.updateTimeframeDropdown();
        this.renderKPIs();
        this.renderChart();
        this.renderTable();
    }

    renderKPIs() {
        const elDO = document.getElementById("dat-kpi-do");
        const elSku = document.getElementById("dat-kpi-sku");
        const elVol = document.getElementById("dat-kpi-vol");
        const elQty = document.getElementById("dat-kpi-qty");
        const elLclDo = document.getElementById("dat-kpi-lcl-do");
        const elLclVol = document.getElementById("dat-kpi-lcl-vol");

        if (this.history.length === 0) {
            const set = (el, val) => { if (el) el.innerText = val; };
            set(elDO, "0"); set(elSku, "0"); set(elVol, "0.00 m³"); set(elQty, "0");
            set(elLclDo, "0"); set(elLclVol, "0.00 m³");
            ["dat-kpi-do-delta", "dat-kpi-sku-delta", "dat-kpi-vol-delta", "dat-kpi-qty-delta"].forEach(id => {
                const el = document.getElementById(id);
                if (el) el.innerText = "No data";
            });
            return;
        }

        // Latest & previous day deltas (matching FormatKpiCardMain)
        const latest = this.history[this.history.length - 1];
        const prev = this.history.length > 1 ? this.history[this.history.length - 2] : null;

        const renderDelta = (cur, prevVal, labelId, fmt) => {
            const el = document.getElementById(labelId);
            if (!el) return;
            if (prevVal === null) {
                el.innerText = "First entry";
                return;
            }
            const absChange = cur - prevVal;
            const pct = prevVal !== 0 ? Math.abs(absChange / prevVal) * 100 : 0;
            const arrow = absChange > 0 ? "▲" : absChange < 0 ? "▼" : "●";
            const sign = absChange > 0 ? "+" : "";
            const color = absChange > 0 ? "#10b981" : absChange < 0 ? "#ef4444" : "#71717a";
            el.innerHTML = `<span style="color:${color}; font-weight:700;">${arrow} ${sign}${fmt(absChange)} (${sign}${pct.toFixed(1)}% vs prev)</span>`;
        };

        const fmtInt = (n) => n.toLocaleString();
        const fmtVol = (n) => n.toFixed(2);

        if (elDO) elDO.innerText = latest.do.toLocaleString();
        if (elSku) elSku.innerText = latest.sku.toLocaleString();
        if (elVol) elVol.innerText = `${latest.vol.toFixed(2)} m³`;
        if (elQty) elQty.innerText = latest.qty.toLocaleString();
        if (elLclDo) elLclDo.innerText = latest.lclDo.toLocaleString();
        if (elLclVol) elLclVol.innerText = `${latest.lclVol.toFixed(2)} m³`;

        renderDelta(latest.do, prev ? prev.do : null, "dat-kpi-do-delta", fmtInt);
        renderDelta(latest.sku, prev ? prev.sku : null, "dat-kpi-sku-delta", fmtInt);
        renderDelta(latest.vol, prev ? prev.vol : null, "dat-kpi-vol-delta", fmtVol);
        renderDelta(latest.qty, prev ? prev.qty : null, "dat-kpi-qty-delta", fmtInt);
    }

    renderChart() {
        const canvas = document.getElementById("activityTrendChart");
        if (!canvas) return;

        const elTitle = document.getElementById("datChartTitle");
        const elSub = document.getElementById("datChartSub");

        if (this.history.length === 0) {
            if (this.chart) { this.chart.destroy(); this.chart = null; }
            return;
        }

        const filtered = this.getFilteredHistory();
        const labels = filtered.map(h => {
            const [y, m, d] = h.date.split("-");
            return `${d}/${m}/${y}`;
        });

        const isLcl = this.chartType === "lcl";

        // Update header labels based on chart type
        if (elTitle) {
            elTitle.innerText = isLcl ? "Daily LCL Activity Trend" : "Daily DO Activity Trend";
        }
        if (elSub) {
            elSub.innerText = isLcl
                ? "LCL QTY columns with DO & Volume lines"
                : "DO & SKU columns with Volume line";
        }

        const isLight = isLightTheme();
        const axisColor = isLight ? "#0f172a" : "#a1a1aa";
        const dateTextColor = isLight ? "#0f172a" : "#e4e4e7";
        const gridColor = isLight ? "#e2e8f0" : "#18181b";

        if (this.chart) this.chart.destroy();

        let datasets;
        if (isLcl) {
            // LCL view: QTY bars (primary), DO bars + Volume line (secondary axis) - mirrors VBA LCL scale fix
            datasets = [
                {
                    label: "LCL Quantity",
                    data: filtered.map(h => h.lclQty),
                    backgroundColor: "#06b6d4",
                    borderRadius: 4,
                    yAxisID: "y"
                },
                {
                    label: "LCL DO Count",
                    data: filtered.map(h => h.lclDo),
                    backgroundColor: "#8b5cf6",
                    borderRadius: 4,
                    yAxisID: "y1"
                },
                {
                    label: "LCL Volume (M3)",
                    data: filtered.map(h => h.lclVol),
                    type: 'line',
                    borderColor: "#f59e0b",
                    backgroundColor: "#f59e0b",
                    pointBackgroundColor: "#f59e0b",
                    pointBorderColor: "#f59e0b",
                    borderWidth: 2,
                    pointRadius: 4,
                    tension: 0.3,
                    yAxisID: "y1"
                }
            ];
        } else {
            datasets = [
                {
                    label: "DO Count",
                    data: filtered.map(h => h.do),
                    backgroundColor: "#10b981",
                    borderRadius: 4,
                    yAxisID: "y"
                },
                {
                    label: "SKUs Managed",
                    data: filtered.map(h => h.sku),
                    backgroundColor: "#8b5cf6",
                    borderRadius: 4,
                    yAxisID: "y"
                },
                {
                    label: "Volume (M3)",
                    data: filtered.map(h => h.vol),
                    type: 'line',
                    borderColor: "#f59e0b",
                    backgroundColor: "#f59e0b",
                    pointBackgroundColor: "#f59e0b",
                    pointBorderColor: "#f59e0b",
                    borderWidth: 2,
                    pointRadius: 4,
                    tension: 0.3,
                    yAxisID: "y1"
                }
            ];
        }

        this.chart = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { labels: { color: axisColor, font: { size: 12, weight: '700' } } }
                },
                scales: {
                    x: {
                        ticks: {
                            color: dateTextColor,
                            maxRotation: 45,
                            minRotation: 45,
                            font: { size: 11, weight: '700' }
                        },
                        grid: { color: gridColor }
                    },
                    y: {
                        type: 'linear', position: 'left', beginAtZero: true,
                        ticks: { color: axisColor, font: { size: 11, weight: '600' } },
                        grid: { color: gridColor },
                        title: { display: true, text: isLcl ? "LCL QTY" : "DO / SKU", color: axisColor, font: { size: 12, weight: '700' } }
                    },
                    y1: {
                        type: 'linear', position: 'right', beginAtZero: true,
                        ticks: { color: axisColor, font: { size: 11, weight: '600' } },
                        grid: { drawOnChartArea: false },
                        title: { display: true, text: "Volume (m³) / DO", color: axisColor, font: { size: 12, weight: '700' } }
                    }
                }
            }
        });
    }

    changePage(page) {
        const timeframeHistory = this.getFilteredHistory();
        const filtered = timeframeHistory.filter(h => {
            if (!this.searchQuery) return true;
            return h.date.toLowerCase().includes(this.searchQuery);
        });
        const totalPages = this.pageSize > 0 ? (Math.ceil(filtered.length / this.pageSize) || 1) : 1;
        if (page < 1 || page > totalPages) return;
        this.currentPage = page;
        this.renderTable();
    }

    changePageSize(size) {
        this.pageSize = parseInt(size, 10) || 10;
        this.currentPage = 1;
        this.renderTable();
    }

    getPaginatedHistory(filtered) {
        if (this.pageSize <= 0) return filtered;
        const totalPages = Math.ceil(filtered.length / this.pageSize) || 1;
        if (this.currentPage > totalPages) this.currentPage = totalPages;
        if (this.currentPage < 1) this.currentPage = 1;
        const startIndex = (this.currentPage - 1) * this.pageSize;
        return filtered.slice(startIndex, startIndex + this.pageSize);
    }

    renderPagination(totalFilteredItems) {
        const paginationContainer = document.getElementById("datPagination");
        if (!paginationContainer) return;

        if (totalFilteredItems === 0) {
            paginationContainer.innerHTML = "";
            paginationContainer.style.display = "none";
            return;
        }

        paginationContainer.style.display = "flex";

        const pageSize = this.pageSize;
        const currentPage = this.currentPage;
        const totalPages = pageSize > 0 ? Math.ceil(totalFilteredItems / pageSize) : 1;

        let startItem = 1;
        let endItem = totalFilteredItems;

        if (pageSize > 0) {
            startItem = (currentPage - 1) * pageSize + 1;
            endItem = Math.min(currentPage * pageSize, totalFilteredItems);
        }

        let paginationHTML = `
            <div class="pagination-info">
                Showing <strong>${startItem}–${endItem}</strong> of <strong>${totalFilteredItems}</strong> day(s)
            </div>
        `;

        if (pageSize > 0 && totalPages > 1) {
            paginationHTML += `<div class="pagination-controls">`;

            paginationHTML += `
                <button class="pagination-btn" onclick="activityTrend.changePage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>
                    ‹ Prev
                </button>
            `;

            let pageNumbers = [];
            const maxPagesToShow = 5;
            let startPage = Math.max(1, currentPage - 2);
            let endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);
            if (endPage - startPage < maxPagesToShow - 1) {
                startPage = Math.max(1, endPage - maxPagesToShow + 1);
            }

            if (startPage > 1) {
                pageNumbers.push(1);
                if (startPage > 2) pageNumbers.push('...');
            }
            for (let p = startPage; p <= endPage; p++) {
                pageNumbers.push(p);
            }
            if (endPage < totalPages) {
                if (endPage < totalPages - 1) pageNumbers.push('...');
                pageNumbers.push(totalPages);
            }

            pageNumbers.forEach(p => {
                if (p === '...') {
                    paginationHTML += `<span class="pagination-ellipsis">...</span>`;
                } else {
                    paginationHTML += `
                        <button class="pagination-btn ${p === currentPage ? 'active' : ''}" onclick="activityTrend.changePage(${p})">
                            ${p}
                        </button>
                    `;
                }
            });

            paginationHTML += `
                <button class="pagination-btn" onclick="activityTrend.changePage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>
                    Next ›
                </button>
            `;

            paginationHTML += `</div>`;
        }

        paginationContainer.innerHTML = paginationHTML;
    }

    renderTable() {
        const tbody = document.getElementById("datTableBody");
        if (!tbody) return;

        const elSearchCount = document.getElementById("datSearchCount");

        if (this.history.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" style="text-align:center; padding:24px; color:#71717a;">No activity data loaded. Please import daily reports.</td></tr>';
            if (elSearchCount) elSearchCount.style.display = "none";
            this.renderPagination(0);
            return;
        }

        const timeframeHistory = this.getFilteredHistory();

        const filtered = timeframeHistory.filter(h => {
            if (!this.searchQuery) return true;
            return h.date.toLowerCase().includes(this.searchQuery);
        });

        if (elSearchCount) {
            elSearchCount.style.display = "inline-block";
            elSearchCount.innerText = this.searchQuery ? `${filtered.length} / ${timeframeHistory.length} days` : `${timeframeHistory.length} days`;
        }

        if (filtered.length === 0) {
            tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:24px; color:#71717a;">No days match search "${this.searchQuery}".</td></tr>`;
            this.renderPagination(0);
            return;
        }

        const paginated = this.getPaginatedHistory(filtered);

        let html = "";
        paginated.forEach(h => {
            const [y, m, d] = h.date.split("-");
            html += `<tr>
                <td class="dat-date-cell">${d}/${m}/${y}</td>
                <td style="text-align: right; font-weight: 600; padding: 8px 12px;">${h.do.toLocaleString()}</td>
                <td style="text-align: right; padding: 8px 12px;">${h.sku.toLocaleString()}</td>
                <td style="text-align: right; padding: 8px 12px;">${h.vol.toFixed(2)}</td>
                <td style="text-align: right; padding: 8px 12px;">${h.qty.toLocaleString()}</td>
                <td style="text-align: right; color: ${isLightTheme() ? '#0E7C86' : '#22d3ee'}; padding: 8px 12px;">${h.lclDo.toLocaleString()}</td>
                <td style="text-align: right; padding: 8px 12px;">${h.lclSku.toLocaleString()}</td>
                <td style="text-align: right; padding: 8px 12px;">${h.lclVol.toFixed(2)}</td>
                <td style="text-align: right; padding: 8px 12px;">${h.lclQty.toLocaleString()}</td>
            </tr>`;
        });

        tbody.innerHTML = html;
        this.renderPagination(filtered.length);
    }

    // Export historical data to Excel workbook (matching the 3 database sheets structure)
    async exportToExcel() {
        if (this.history.length === 0) {
            alert("[WARNING] No activity history to export! Please import daily reports first.");
            return;
        }

        // Guarantee chronological date sorting (earliest to latest) before building worksheets
        this.history.sort((a, b) => a.date.localeCompare(b.date));

        const workbook = XLSX.utils.book_new();

        // Sheet 1: Summary_Data (Date, DO, SKU, Vol, Qty)
        const summaryData = [["Date", "DO", "SKU", "VOL (m3)", "QTY"]];
        this.history.forEach(h => {
            summaryData.push([h.date, h.do, h.sku, h.vol, h.qty]);
        });
        const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
        wsSummary['!cols'] = [{ wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 8 }];
        XLSX.utils.book_append_sheet(workbook, wsSummary, "Summary_Data");

        // Sheet 2: LCL_Historical_Data (Only include dates with active LCL activity)
        const lclData = [["Date", "LCL DO", "LCL SKU", "LCL VOL", "LCL QTY"]];
        this.history.forEach(h => {
            if (h.lclDo > 0 || h.lclQty > 0 || h.lclVol > 0) {
                lclData.push([h.date, h.lclDo, h.lclSku, h.lclVol, h.lclQty]);
            }
        });
        const wsLcl = XLSX.utils.aoa_to_sheet(lclData);
        wsLcl['!cols'] = [{ wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 8 }];
        XLSX.utils.book_append_sheet(workbook, wsLcl, "LCL_Historical_Data");

        // Determine file name (ddmmyyyy of today)
        const now = new Date();
        const dd = String(now.getDate()).padStart(2, '0');
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const yyyy = now.getFullYear();
        const exportFileName = `DO_Activity_Trend_${dd}${mm}${yyyy}.xlsx`;

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
let activityTrend = null;
document.addEventListener("DOMContentLoaded", () => {
    activityTrend = new DOActivityTrend();
});
