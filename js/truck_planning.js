// Truck Planning & Daily Summary List Controller

let TruckPlanningRawDOs = [];
let TruckPlanningLookupMap = {};

let SelectedTruckFilter = "ALL";
let SelectedTypeFilter = "ALL";
let SelectedHubFilter = "ALL";
let SelectedCategoryFilter = "ALL";
let TruckSearchTerm = "";

// Initialize on page load
window.addEventListener("DOMContentLoaded", () => {
    if (typeof initTheme === 'function') initTheme();
    loadTruckPlanningData();
    updateDateDisplay();
});

// Update top date display (matches Excel 31/7/2026 timestamp format)
function updateDateDisplay() {
    const dateElem = document.getElementById("currentDateDisplay");
    if (!dateElem) return;
    const now = new Date();
    const formatted = `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    dateElem.innerText = formatted;
}

// Load dataset from localStorage
function loadTruckPlanningData() {
    const rawDoData = localStorage.getItem("LastUploadedDoSummary");
    const rawRouteData = localStorage.getItem("LastUploadedRouteData");

    if (rawRouteData) {
        try { TruckPlanningLookupMap = JSON.parse(rawRouteData); } catch(e) {}
    }

    if (rawDoData) {
        try { TruckPlanningRawDOs = JSON.parse(rawDoData); } catch(e) {}
    }

    populateSlicerButtons();
    renderTruckPlanningDashboard();
}

// Extract unique values and build interactive slicer button grids
// Truck and Hub slicers cross-filter: selecting a truck limits hubs to that
// truck (and vice-versa), so "All Hubs" after picking a truck = that truck's hubs.
function populateSlicerButtons() {
    const trucksSet = new Set();
    const hubsSet = new Set();
    const categoriesSet = new Set();

    // Iterate through all records to harvest unique trucks, hubs, and product categories
    TruckPlanningRawDOs.forEach(doRow => {
        const routeMatch = TruckPlanningLookupMap[doRow.inv] || {};
        const truck = routeMatch.truck ? String(routeMatch.truck) : null;
        const hub = routeMatch.hub ? String(routeMatch.hub) : null;

        if (truck) trucksSet.add(truck);
        if (hub) hubsSet.add(hub);

        if (routeMatch.items && routeMatch.items.length > 0) {
            routeMatch.items.forEach(item => {
                const desc = item.desc || "";
                if (desc.includes("(")) {
                    const catMatch = desc.match(/\(([^)]+)\)/);
                    if (catMatch) categoriesSet.add(catMatch[1].toUpperCase());
                }
            });
        }
    });

    // Cross-filtered subsets
    const trucksForHubSet = new Set();
    const hubsForTruckSet = new Set();
    TruckPlanningRawDOs.forEach(doRow => {
        const routeMatch = TruckPlanningLookupMap[doRow.inv] || {};
        const truck = routeMatch.truck ? String(routeMatch.truck) : null;
        const hub = routeMatch.hub ? String(routeMatch.hub) : null;
        if (!truck || !hub) return;
        if (SelectedHubFilter === "ALL" || hub === SelectedHubFilter) trucksForHubSet.add(truck);
        if (SelectedTruckFilter === "ALL" || truck === SelectedTruckFilter) hubsForTruckSet.add(hub);
    });

    // 1. Render Truck Slicer Grid
    const truckGrid = document.getElementById("truckSlicerGrid");
    if (truckGrid) {
        const sortedTrucks = Array.from(trucksSet).sort((a, b) => parseInt(a) - parseInt(b));
        let html = `<button class="slicer-btn ${SelectedTruckFilter === 'ALL' ? 'active' : ''}" onclick="setTruckFilter('ALL')">All Trucks</button>`;
        sortedTrucks.forEach(tr => {
            const disabled = !trucksForHubSet.has(tr);
            html += `<button class="slicer-btn ${SelectedTruckFilter === tr ? 'active' : ''}" ${disabled ? 'disabled style="opacity:0.35;cursor:not-allowed;"' : `onclick="setTruckFilter('${tr}')"`}>Truck ${tr}</button>`;
        });
        truckGrid.innerHTML = html;
    }

    // 2. Render Hub Slicer Grid (limited to selected truck's hubs)
    const hubGrid = document.getElementById("hubSlicerGrid");
    if (hubGrid) {
        const sortedHubs = Array.from(hubsSet).sort();
        let html = `<button class="slicer-btn ${SelectedHubFilter === 'ALL' ? 'active' : ''}" onclick="setHubFilter('ALL')">All Hubs</button>`;
        sortedHubs.forEach(hub => {
            const disabled = !hubsForTruckSet.has(hub);
            html += `<button class="slicer-btn ${SelectedHubFilter === hub ? 'active' : ''}" ${disabled ? 'disabled style="opacity:0.35;cursor:not-allowed;"' : `onclick="setHubFilter('${hub}')"`}>${hub}</button>`;
        });
        hubGrid.innerHTML = html;
    }

    // 3. Render Category Slicer Grid
    const catGrid = document.getElementById("categorySlicerGrid");
    if (catGrid) {
        const sortedCats = Array.from(categoriesSet).sort();
        let html = `<button class="slicer-btn ${SelectedCategoryFilter === 'ALL' ? 'active' : ''}" onclick="setCategoryFilter('ALL')">All Categories</button>`;
        sortedCats.forEach(cat => {
            html += `<button class="slicer-btn ${SelectedCategoryFilter === cat ? 'active' : ''}" onclick="setCategoryFilter('${cat}')">${cat}</button>`;
        });
        catGrid.innerHTML = html;
    }
}

// Slicer Button Click Handlers
function setTruckFilter(val) {
    SelectedTruckFilter = val;
    populateSlicerButtons();
    renderTruckPlanningDashboard();
}

function setTypeFilter(val) {
    SelectedTypeFilter = val;
    // update type slicer active states
    document.querySelectorAll("#typeSlicerGrid .slicer-btn").forEach(btn => {
        btn.classList.toggle("active", btn.getAttribute("data-type") === val);
    });
    renderTruckPlanningDashboard();
}

function setHubFilter(val) {
    SelectedHubFilter = val;
    populateSlicerButtons();
    renderTruckPlanningDashboard();
}

function setCategoryFilter(val) {
    SelectedCategoryFilter = val;
    populateSlicerButtons();
    renderTruckPlanningDashboard();
}

function clearAllFilters() {
    SelectedTruckFilter = "ALL";
    SelectedTypeFilter = "ALL";
    SelectedHubFilter = "ALL";
    SelectedCategoryFilter = "ALL";
    TruckSearchTerm = "";
    const searchInput = document.getElementById("truckSearchInput");
    if (searchInput) searchInput.value = "";
    populateSlicerButtons();
    renderTruckPlanningDashboard();
}

// Filter and Render Main Manifest Table + Update KPI Boxes
function renderTruckPlanningDashboard() {
    TruckSearchTerm = (document.getElementById("truckSearchInput")?.value || "").toLowerCase().trim();

    let filteredRows = [];

    TruckPlanningRawDOs.forEach(doRow => {
        const routeMatch = TruckPlanningLookupMap[doRow.inv] || {};
        const truckNo = String(routeMatch.truck || "N/A");
        const hubId = String(routeMatch.hub || "N/A");
        const doType = (routeMatch.doCategory || routeMatch.doType || "BIG").toUpperCase();
        const remark = doRow.remark || "";
        const fullDoText = remark ? `${doRow.inv} ${remark}` : doRow.inv;

        // Apply Slicers
        if (SelectedTruckFilter !== "ALL" && truckNo !== SelectedTruckFilter) return;
        if (SelectedTypeFilter !== "ALL" && doType !== SelectedTypeFilter.toUpperCase()) return;
        if (SelectedHubFilter !== "ALL" && hubId !== SelectedHubFilter) return;

        const itemsList = (routeMatch.items && routeMatch.items.length > 0) 
            ? routeMatch.items 
            : [{ code: "Unspecified", desc: "No Product Description", qty: doRow.qty }];

        itemsList.forEach(item => {
            const code = item.code || "-";
            const desc = item.desc || "-";
            const qty = item.qty || 0;

            // Category Filter Check
            if (SelectedCategoryFilter !== "ALL") {
                if (!desc.toUpperCase().includes(`(${SelectedCategoryFilter})`)) return;
            }

            // Search Term Filter Check
            if (TruckSearchTerm) {
                const searchable = `${doRow.inv} ${doRow.route} ${doRow.name} ${doRow.addr} ${truckNo} ${hubId} ${code} ${desc} ${remark}`.toLowerCase();
                if (!searchable.includes(TruckSearchTerm)) return;
            }

            filteredRows.push({
                truck: truckNo,
                hub: hubId,
                route: doRow.route || "-",
                consignee: doRow.name || "-",
                doNo: fullDoText,
                invRaw: doRow.inv,
                type: doType,
                productCode: code,
                modelName: desc,
                qty: qty,
                vol: parseFloat(doRow.vol) || 0
            });
        });
    });

    // Calculate KPI Totals
    const uniqueDos = new Set(filteredRows.map(r => r.invRaw)).size;
    const uniqueModels = new Set(filteredRows.map(r => r.productCode)).size;
    const totalQty = filteredRows.reduce((sum, r) => sum + r.qty, 0);

    // Total m3 summed over UNIQUE DOs (vol is per-DO, repeats across item rows)
    const volByDo = {};
    filteredRows.forEach(r => { volByDo[r.invRaw] = r.vol; });
    const totalM3 = Object.values(volByDo).reduce((sum, v) => sum + v, 0);

    document.getElementById("kpiTotalDos").innerText = uniqueDos.toLocaleString();
    document.getElementById("kpiTotalModels").innerText = uniqueModels.toLocaleString();
    document.getElementById("kpiTotalQty").innerText = totalQty.toLocaleString();
    const m3Kpi = document.getElementById("kpiTotalM3");
    if (m3Kpi) m3Kpi.innerText = totalM3.toFixed(2) + " m³";

    // Render Table Rows
    // When a Truck is selected, hide Truck/Hub/Route/Consignee columns and show a
    // manifest header (TRUCK / HUB / CONSIGNEE / ROUTE) above the table. Selecting
    // a truck auto-filters hubs to that truck, so "All Hubs" = that truck's hubs.
    const tbody = document.getElementById("truckPlanningTableBody");
    const headRow = document.getElementById("truckPlanningTableHeadRow");
    const manifestHeader = document.getElementById("truckManifestHeader");
    if (!tbody) return;

    const showManifestView = SelectedTruckFilter !== "ALL";

    if (filteredRows.length === 0) {
        if (manifestHeader) manifestHeader.style.display = "none";
        tbody.innerHTML = `<tr><td colspan="${showManifestView ? 6 : 10}" style="text-align:center; padding:32px; color:#a1a1aa;">No DO manifests match current filter selection.</td></tr>`;
        return;
    }

    let tableHtml = "";

    if (showManifestView) {
        // Populate manifest header from filtered rows (dedupe hubs/consignees/routes)
        const uniqueConsignees = [...new Set(filteredRows.map(r => r.consignee))].join(", ");
        const uniqueRoutes = [...new Set(filteredRows.map(r => r.route))].join(", ");
        const uniqueHubs = [...new Set(filteredRows.map(r => r.hub))].join(", ");

        if (manifestHeader) {
            manifestHeader.style.display = "grid";
            document.getElementById("manifestTruck").innerText = filteredRows[0].truck;
            document.getElementById("manifestConsignee").innerText = uniqueConsignees;
            document.getElementById("manifestHub").innerText = uniqueHubs;
            document.getElementById("manifestRoute").innerText = uniqueRoutes;
        }

        if (headRow) {
            headRow.innerHTML = `
                <th>DO Number</th>
                <th>Type</th>
                <th>Product Code</th>
                <th>Model Name</th>
                <th class="number-col">Total</th>
                <th class="number-col">m³</th>`;
        }

        filteredRows.forEach(row => {
            const typeBadge = row.type === 'BIG' ? 'category-big' : row.type === 'SMALL' ? 'category-small' : 'category-mix';
            tableHtml += `
            <tr style="cursor: context-menu;" title="Right-click to inspect DO details" oncontextmenu="event.preventDefault(); openDoDetailsTab('${row.invRaw}')">
                <td><strong>${row.doNo}</strong></td>
                <td><span class="category-badge ${typeBadge}">${row.type}</span></td>
                <td style="font-family: monospace; font-weight: 700; color: #10b981;">${row.productCode}</td>
                <td>${row.modelName}</td>
                <td class="number-col" style="font-weight: 800; color: #8b5cf6;">${row.qty.toLocaleString()}</td>
                <td class="number-col" style="font-weight: 700; color: #f59e0b;">${row.vol.toFixed(2)}</td>
            </tr>
            `;
        });
    } else {
        if (manifestHeader) manifestHeader.style.display = "none";
        if (headRow) {
            headRow.innerHTML = `
                <th>Truck</th>
                <th>Hub</th>
                <th>Route</th>
                <th>Consignee</th>
                <th>DO Number</th>
                <th>Type</th>
                <th>Product Code</th>
                <th>Model Name</th>
                <th class="number-col">Total</th>
                <th class="number-col">m³</th>`;
        }

        filteredRows.forEach(row => {
            const typeBadge = row.type === 'BIG' ? 'category-big' : row.type === 'SMALL' ? 'category-small' : 'category-mix';
            tableHtml += `
            <tr style="cursor: context-menu;" title="Right-click to inspect DO details" oncontextmenu="event.preventDefault(); openDoDetailsTab('${row.invRaw}')">
                <td><strong style="color: #3b82f6;">${row.truck}</strong></td>
                <td><strong>${row.hub}</strong></td>
                <td>${row.route}</td>
                <td style="max-width: 180px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${row.consignee}</td>
                <td><strong>${row.doNo}</strong></td>
                <td><span class="category-badge ${typeBadge}">${row.type}</span></td>
                <td style="font-family: monospace; font-weight: 700; color: #10b981;">${row.productCode}</td>
                <td>${row.modelName}</td>
                <td class="number-col" style="font-weight: 800; color: #8b5cf6;">${row.qty.toLocaleString()}</td>
                <td class="number-col" style="font-weight: 700; color: #f59e0b;">${row.vol.toFixed(2)}</td>
            </tr>
            `;
        });
    }
    tbody.innerHTML = tableHtml;
}


// Opens DO breakdown detail page in a new tab upon right-click
function openDoDetailsTab(invoiceNo) {
    if (!invoiceNo) return;
    localStorage.setItem("SelectedDoForDetails", invoiceNo);
    localStorage.setItem("DoDetailsReturnTo", "truck_planning.html");
    window.open('do_details.html', '_blank');
}
