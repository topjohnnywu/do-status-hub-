// Update memory status badge with clean chip pills (file names only)
function updateMemoryBadge() {
    const BadgeElement = document.getElementById("memoryStatusBadge");
    if (!BadgeElement) return;

    const SavedDoName = localStorage.getItem("LastDoSummaryFileName");
    const SavedRouteName = localStorage.getItem("LastRouteFileName");
    const SavedInsightName = localStorage.getItem("LastShippingInsightFileName");

    if (!SavedDoName && !SavedRouteName && !SavedInsightName) {
        BadgeElement.innerHTML = "";
        return;
    }

    let html = "";

    if (SavedDoName) {
        html += `<span class="file-chip green">${SavedDoName}</span>`;
    }
    if (SavedRouteName) {
        html += `<span class="file-chip blue">${SavedRouteName}</span>`;
    }
    if (SavedInsightName) {
        html += `<span class="file-chip purple">${SavedInsightName}</span>`;
    }

    BadgeElement.innerHTML = html;
}

// Smart Multi-File Excel Parser with Column L (Remarks) Extraction
async function handleFileUpload(event) {
    const FileListObjects = Array.from(event.target.files);
    if (FileListObjects.length === 0) return;

    MasterFileStoreArray = [];
    const FileSelectorDropdown = document.getElementById("fileSelector");
    if (FileSelectorDropdown) {
        FileSelectorDropdown.innerHTML = `<option value="ALL">All Files Combined (${FileListObjects.length})</option>`;
    }

    for (const FilePickerElement of FileListObjects) {
        if (FileSelectorDropdown) {
            const OptionElement = document.createElement("option");
            OptionElement.value = FilePickerElement.name;
            OptionElement.textContent = FilePickerElement.name;
            FileSelectorDropdown.appendChild(OptionElement);
        }

        await new Promise((resolve) => {
            const ReaderObject = new FileReader();
            ReaderObject.onload = function(e) {
                try {
                    const RawArrayBuffer = e.target.result;
                    const WorkbookObject = XLSX.read(RawArrayBuffer, { type: 'array' });

                    let SelectedSheetName = WorkbookObject.SheetNames.find(name => {
                        const lower = name.toLowerCase();
                        return lower.includes("summary") || lower.includes("do") || lower.includes("final");
                    }) || WorkbookObject.SheetNames[0];

                    const TargetSheetHoarder = WorkbookObject.Sheets[SelectedSheetName];

                    if (TargetSheetHoarder) {
                        const RawRows = XLSX.utils.sheet_to_json(TargetSheetHoarder, { header: 1, raw: false });

                        let DataStartRowIndex = 1;
                        let TrailingColIdx = 6;
                        let LeadingColIdx = 7;

                        for (let r = 0; r < Math.min(5, RawRows.length); r++) {
                            const rowText = (RawRows[r] || []).map(cell => String(cell || "").toLowerCase().trim());
                            
                            const firstCell = String(RawRows[r]?.[0] || "").toUpperCase();
                            if (firstCell.includes("INVOICE") || firstCell.includes("DO") || firstCell.includes("SHIPMENT")) {
                                DataStartRowIndex = r + 1;
                            }

                            const tIdx = rowText.findIndex(t => t === "trailing" || t === "status 1");
                            if (tIdx !== -1) TrailingColIdx = tIdx;

                            const lIdx = rowText.findIndex(t => t === "leading status" || t === "leading");
                            if (lIdx !== -1) LeadingColIdx = lIdx;
                        }

                        for (let HamsterWheelIndex = DataStartRowIndex; HamsterWheelIndex < RawRows.length; HamsterWheelIndex++) {
                            const row = RawRows[HamsterWheelIndex];
                            if (!row || !row[0] || String(row[0]).toUpperCase().includes("TOTAL") || String(row[0]).toUpperCase() === "SHIPMENT") continue;

                            MasterFileStoreArray.push({
                                fileName: FilePickerElement.name,
                                inv: String(row[0] || "").trim(),
                                div: String(row[1] || "N/A").trim(),
                                code: String(row[2] || "").trim(),
                                route: String(row[3] || "").trim(),
                                name: String(row[4] || "").trim(),
                                addr: [row[5]].filter(Boolean).join(" ").trim(),
                                colG: String(row[TrailingColIdx] || "Missing").trim(),
                                colH: String(row[LeadingColIdx] || "Missing").trim(),
                                vol: parseFloat(row[8]) || 0,
                                qty: parseInt(row[9]) || 0,
                                sku: parseInt(row[10]) || 0,
                                remark: String(row[11] || "").trim() // Column L (Remarks)
                            });
                        }
                    }
                } catch (err) {
                    console.error("File parsing error:", err);
                }
                resolve();
            };
            ReaderObject.readAsArrayBuffer(FilePickerElement);
        });
    }

    if (MasterFileStoreArray.length === 0) {
        alert("[WARNING] Could not extract data rows from uploaded file(s)!");
    }

    localStorage.setItem("LastUploadedDoSummary", JSON.stringify(MasterFileStoreArray));
    const DoNames = FileListObjects.map(f => f.name).join(", ");
    localStorage.setItem("LastDoSummaryFileName", DoNames);

    if (typeof updateMemoryBadge === 'function') updateMemoryBadge();

    DataHoarderArray = [...MasterFileStoreArray];
    refreshDashboard();
}

// Process Planning / Source Files (.xlsm / .xlsx / .csv) targeting 'Insert Batch' sheet tab
async function handleProductMasterUpload(event) {
    const SourceFiles = Array.from(event.target.files);
    if (SourceFiles.length === 0) return;

    for (const SourceFile of SourceFiles) {
        await new Promise((resolve) => {
            const ReaderObject = new FileReader();
            const IsCsvFile = SourceFile.name.toLowerCase().endsWith('.csv');

            ReaderObject.onload = function(e) {
                let RawRows = [];

                if (IsCsvFile) {
                    const CsvText = e.target.result;
                    const WorkbookObject = XLSX.read(CsvText, { type: 'string' });
                    RawRows = XLSX.utils.sheet_to_json(WorkbookObject.Sheets[WorkbookObject.SheetNames[0]], { header: 1 });
                } else {
                    const RawArrayBuffer = e.target.result;
                    const WorkbookObject = XLSX.read(RawArrayBuffer, { type: 'array' });

                    // Target the 'Insert Batch' sheet tab specifically
                    let TargetSheetName = WorkbookObject.SheetNames.find(name => {
                        const lower = name.toLowerCase().trim();
                        return lower === "insert batch" || lower.includes("insert batch") || lower.includes("batch");
                    }) || WorkbookObject.SheetNames[0];

                    const TargetSheet = WorkbookObject.Sheets[TargetSheetName];
                    RawRows = XLSX.utils.sheet_to_json(TargetSheet, { header: 1 });
                }

                if (RawRows.length < 4) { resolve(); return; }

                // Data starts at Row 5 (Index 4), header is at Row 4 (Index 3)
                for (let i = 4; i < RawRows.length; i++) {
                    const row = RawRows[i];
                    if (!row || !row[3]) continue; // Skip empty DO rows

                    const RawDoStr = String(row[3]).trim(); // Col D (Index 3) = DO Number
                    if (!RawDoStr || RawDoStr.toUpperCase().includes("TOTAL") || RawDoStr.toUpperCase().includes("HUB")) continue;

                    // Extract full DO number (supports 8-digit, 10-digit, or any digit length)
                    const CleanedFirstToken = RawDoStr.split(" ")[0].trim();
                    const DigitMatch = RawDoStr.match(/\d+/);
                    const InvoiceKey = DigitMatch ? DigitMatch[0] : CleanedFirstToken;

                    const TruckVal = String(row[1] || "").trim(); // Col B (Index 1) = Truck Number
                    const HubVal = String(row[2] || "").trim();   // Col C (Index 2) = Hub
                    const RouteVal = String(row[4] || "").trim(); // Col E (Index 4) = Route

                    const ColFVal = String(row[5] || "").trim(); // Col F (Index 5) = Product Code
                    const ColGVal = String(row[6] || "").trim(); // Col G (Index 6) = Ship Qty
                    const ColJVal = String(row[9] || "").trim(); // Col J (Index 9) = Type (Small/Big/Mix)
                    const ColKVal = String(row[10] || "").trim(); // Col K (Index 10) = Model Name

                    // Clean and parse Ship Quantity from Column G
                    const CleanQtyNum = parseInt(ColGVal.replace(/[^0-9]/g, ""), 10);
                    const LineQty = (!isNaN(CleanQtyNum) && CleanQtyNum > 0) ? CleanQtyNum : 1;

                    // Classify item row category directly from Column J (Type)
                    const LowerType = ColJVal.toLowerCase();
                    let RowCategory = "small";
                    if (LowerType.includes("big")) RowCategory = "big";
                    else if (LowerType.includes("mix")) RowCategory = "mix";
                    else if (LowerType.includes("small")) RowCategory = "small";

                    if (!ProductMasterLookupMap[InvoiceKey]) {
                        ProductMasterLookupMap[InvoiceKey] = {
                            items: [], listK: [], listL: [], sizesSet: new Set(), doCategory: RowCategory,
                            truck: TruckVal, hub: HubVal, route: RouteVal
                        };
                    } else {
                        if (TruckVal && !ProductMasterLookupMap[InvoiceKey].truck) ProductMasterLookupMap[InvoiceKey].truck = TruckVal;
                        if (HubVal && !ProductMasterLookupMap[InvoiceKey].hub) ProductMasterLookupMap[InvoiceKey].hub = HubVal;
                        if (RouteVal && !ProductMasterLookupMap[InvoiceKey].route) ProductMasterLookupMap[InvoiceKey].route = RouteVal;
                    }

                    // Save Product Code, Model Name, Ship Quantity, and Route
                    ProductMasterLookupMap[InvoiceKey].items.push({
                        code: ColFVal || "Unspecified Code",
                        desc: ColKVal || "No Description",
                        qty: LineQty,
                        route: RouteVal
                    });

                    if (ColFVal && !ProductMasterLookupMap[InvoiceKey].listK.includes(ColFVal)) {
                        ProductMasterLookupMap[InvoiceKey].listK.push(ColFVal);
                    }
                    if (ColKVal && !ProductMasterLookupMap[InvoiceKey].listL.includes(ColKVal)) {
                        ProductMasterLookupMap[InvoiceKey].listL.push(ColKVal);
                    }

                    // Store row category into the DO's size set
                    ProductMasterLookupMap[InvoiceKey].sizesSet.add(RowCategory);
                }
                resolve();
            };

            if (IsCsvFile) ReaderObject.readAsText(SourceFile);
            else ReaderObject.readAsArrayBuffer(SourceFile);
        });
    }

    // Compute FINAL DO Category across all item rows for each DO
    Object.keys(ProductMasterLookupMap).forEach(key => {
        const SizeSet = ProductMasterLookupMap[key].sizesSet;

        if (SizeSet.has("mix") || (SizeSet.has("big") && SizeSet.has("small"))) {
            ProductMasterLookupMap[key].doCategory = "mix";
        } else if (SizeSet.has("big")) {
            ProductMasterLookupMap[key].doCategory = "big";
        } else {
            ProductMasterLookupMap[key].doCategory = "small";
        }
    });

    // Check for unmatched DOs between Batch Picking and DO Summary
    let UnmatchedMsg = "";
    if (typeof MasterFileStoreArray !== 'undefined' && MasterFileStoreArray.length > 0) {
        const SummarySet = new Set(MasterFileStoreArray.map(item => item.inv));
        const MissingInSummary = Object.keys(ProductMasterLookupMap).filter(key => !SummarySet.has(key));

        if (MissingInSummary.length > 0) {
            const SampleList = MissingInSummary.slice(0, 4).join(", ");
            const OverflowCount = MissingInSummary.length > 4 ? ` (+${MissingInSummary.length - 4} more)` : "";
            UnmatchedMsg = `\n\n[NOTICE] Found ${MissingInSummary.length} DO(s) in Batch Picking missing from DO Summary file: ${SampleList}${OverflowCount}`;
        }
    }

    alert(`[SUCCESS] Processed ${SourceFiles.length} Batch Picking file(s) from 'Insert Batch' sheet!${UnmatchedMsg}`);

    const RouteNames = SourceFiles.map(f => f.name).join(", ");
    localStorage.setItem("LastRouteFileName", RouteNames);
    localStorage.setItem("LastUploadedRouteData", JSON.stringify(ProductMasterLookupMap));

    if (typeof updateMemoryBadge === 'function') updateMemoryBadge();
    refreshDashboard();
}

// Save Master Catalog rules to localStorage
function handleRulesUpload(event) {
    const RulesFile = event.target.files[0];
    if (!RulesFile) return;

    const ReaderObject = new FileReader();
    ReaderObject.onload = function(e) {
        try {
            const RawArrayBuffer = e.target.result;
            const WorkbookObject = XLSX.read(RawArrayBuffer, { type: 'array' });
            const FirstSheet = WorkbookObject.Sheets[WorkbookObject.SheetNames[0]];
            const RawRows = XLSX.utils.sheet_to_json(FirstSheet, { header: 1, raw: false });

            ProductSizeRuleMap = JSON.parse(localStorage.getItem("ProductSizeVault")) || {};
            let SavedCount = 0;

            const firstRowColB = String(RawRows[0]?.[1] || "").toLowerCase().trim();
            const startIdx = (firstRowColB.includes("big") || firstRowColB.includes("small")) ? 0 : 1;

            for (let i = startIdx; i < RawRows.length; i++) {
                const row = RawRows[i];
                if (!row || !row[0]) continue;

                const ProdCode = String(row[0]).trim();
                const RawSize = String(row[1] || "").toLowerCase().trim();

                let SizeVal = "";
                if (RawSize.includes("big")) SizeVal = "big";
                else if (RawSize.includes("small")) SizeVal = "small";

                if (ProdCode && SizeVal) {
                    ProductSizeRuleMap[ProdCode] = SizeVal;
                    SavedCount++;
                }
            }

            localStorage.setItem("ProductSizeVault", JSON.stringify(ProductSizeRuleMap));
            updateRulesStatusUI();
            alert(`[SUCCESS] Saved ${SavedCount.toLocaleString()} product rules! Total Vault: ${Object.keys(ProductSizeRuleMap).length.toLocaleString()} items.`);
            refreshDashboard();

        } catch (err) {
            alert("[ERROR] Failed to parse Master Catalog file: " + err.message);
        }
    };
    ReaderObject.readAsArrayBuffer(RulesFile);
}

// Separate Shipping Insight File Loader with Full Summary Reconciliation
async function handleShippingInsightUpload(event) {
    const InsightFile = event.target.files[0];
    if (!InsightFile) return;

    const ReaderObject = new FileReader();
    ReaderObject.onload = function(e) {
        try {
            const RawArrayBuffer = e.target.result;
            const WorkbookObject = XLSX.read(RawArrayBuffer, { type: 'array' });

            let TargetSheetName = WorkbookObject.SheetNames.find(name => {
                const lower = name.toLowerCase();
                return lower.includes("do summary") || lower.includes("summary") || lower.includes("list") || lower.includes("insight");
            }) || WorkbookObject.SheetNames[0];

            const TargetSheet = WorkbookObject.Sheets[TargetSheetName];
            const RawRows = XLSX.utils.sheet_to_json(TargetSheet, { header: 1, raw: false });

            let ShipmentColIdx = 0, TrailingColIdx = 6, LeadingColIdx = 7, DataStartRow = 1;

            for (let r = 0; r < Math.min(5, RawRows.length); r++) {
                const rowText = (RawRows[r] || []).map(cell => String(cell || "").toLowerCase().trim());
                
                const sIdx = rowText.findIndex(t => t === "shipment" || t === "do number" || t === "invoice no");
                if (sIdx !== -1) ShipmentColIdx = sIdx;

                const tIdx = rowText.findIndex(t => t === "trailing" || t === "status 1");
                if (tIdx !== -1) TrailingColIdx = tIdx;

                const lIdx = rowText.findIndex(t => t === "leading status" || t === "leading");
                if (lIdx !== -1) LeadingColIdx = lIdx;

                if (sIdx !== -1 || tIdx !== -1 || lIdx !== -1) { DataStartRow = r + 1; break; }
            }

            // Fallback to local memory if main summary array is not in active scope
            let DoSummaryList = (typeof MasterFileStoreArray !== 'undefined' && MasterFileStoreArray.length > 0)
                ? MasterFileStoreArray 
                : (JSON.parse(localStorage.getItem("LastUploadedDoSummary")) || []);

            const SummaryDoMap = new Map(DoSummaryList.map(item => [item.inv, item]));
            const SummaryDoSet = new Set(DoSummaryList.map(item => item.inv));

            let MasterLookup = (typeof ProductMasterLookupMap !== 'undefined' && Object.keys(ProductMasterLookupMap).length > 0)
                ? ProductMasterLookupMap 
                : (JSON.parse(localStorage.getItem("LastUploadedRouteData")) || {});

            const PayloadRecords = [];
            const ProcessedDoSet = new Set();

            // 1. Process all scanned DOs from the Shipping Insight file
            for (let i = DataStartRow; i < RawRows.length; i++) {
                const row = RawRows[i];
                if (!row) continue;

                const RawInvoiceVal = String(row[ShipmentColIdx] || "").trim();
                if (!RawInvoiceVal || RawInvoiceVal.toUpperCase().includes("TOTAL") || RawInvoiceVal.toUpperCase() === "SHIPMENT") continue;

                const colGVal = String(row[TrailingColIdx] || "Missing").trim();
                const colHVal = String(row[LeadingColIdx] || "Missing").trim();

                const isCompletedScan = (colGVal.toLowerCase() === "ship confirm pending" && colHVal.toLowerCase() === "ship confirm pending");
                const existsInSummary = SummaryDoSet.has(RawInvoiceVal);

                // Skip unmatched DOs that are NOT completed
                if (!existsInSummary && !isCompletedScan) {
                    continue;
                }

                let computedType = "PENDING";
                let truckVal = "N/A";
                let hubVal = "N/A";

                if (!existsInSummary && isCompletedScan) {
                    computedType = "UNMATCHED DO";
                } else {
                    const Match = MasterLookup[RawInvoiceVal] || {};
                    computedType = (Match.doCategory || "pending").toUpperCase();
                    truckVal = Match.truck || "N/A";
                    hubVal = Match.hub || "N/A";
                }

                const summaryRecord = SummaryDoMap.get(RawInvoiceVal);
                const RouteVal = summaryRecord ? summaryRecord.route : "";
                const ShipToName = summaryRecord ? summaryRecord.name : String(row[9] || "").trim();
                const CombinedDestination = [RouteVal, ShipToName].filter(Boolean).join(" - ");

                PayloadRecords.push({
                    inv: RawInvoiceVal,
                    truck: truckVal,
                    hub: hubVal,
                    addr: CombinedDestination || ShipToName || "N/A",
                    colG: colGVal,
                    colH: colHVal,
                    doType: computedType
                });

                ProcessedDoSet.add(RawInvoiceVal);
            }

            // 2. Reconcile with DO Summary List: Add missing unscanned DOs as "Pending Scan"
            DoSummaryList.forEach(sumItem => {
                if (!ProcessedDoSet.has(sumItem.inv)) {
                    const Match = MasterLookup[sumItem.inv] || {};
                    const catType = (Match.doCategory || "pending").toUpperCase();
                    const CombinedDestination = [sumItem.route, sumItem.name].filter(Boolean).join(" - ");

                    PayloadRecords.push({
                        inv: sumItem.inv,
                        truck: Match.truck || "N/A",
                        hub: Match.hub || "N/A",
                        addr: CombinedDestination || sumItem.name || sumItem.addr || "N/A",
                        colG: "picking pending",
                        colH: "picking pending",
                        doType: catType
                    });
                }
            });

localStorage.setItem("ShippingInsightData", JSON.stringify(PayloadRecords));
localStorage.setItem("LastShippingInsightFileName", InsightFile.name);
if (typeof updateMemoryBadge === 'function') updateMemoryBadge();

            // If uploaded directly inside shipping_insight.html, refresh in-place
            if (typeof applyInsightFilter === 'function') {
                RawPayload = PayloadRecords;
                applyInsightFilter();
                alert(`[SUCCESS] Refreshed dashboard with ${PayloadRecords.length} DO records!`);
            } else {
                // Uploaded from summary.html: Automatically opens shipping_insight.html in a NEW TAB
                alert(`[SUCCESS] Reconciled all ${PayloadRecords.length} DOs! Opening Shipping Insight in a new tab...`);
                window.open('shipping_insight.html', '_blank');
            }

        } catch (err) {
            console.error("Shipping Insight Error Details:", err);
            alert("[ERROR] Failed to parse Shipping Insight file: " + err.message);
        }
    };

    ReaderObject.readAsArrayBuffer(InsightFile);
}