// Chart Handles
let TopConsigneeChart = null;

const sunIconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#f59e0b" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>`;
const moonIconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#a78bfa" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>`;

function updateThemeToggleButton(themeName) {
    const ToggleBtn = document.getElementById('themeToggleBtn');
    if (!ToggleBtn) return;
    const isLight = themeName === "light" || themeName === "organic" || themeName === "mono" || themeName === "bauhaus" || (themeName === "github" && document.body.classList.contains('github-light'));
    const label = ({ linear: "Linear", terminal: "Terminal", dopamine: "Dopamine", amoled: "AMOLED", light: "Light", retro: "Retro", organic: "Organic", cyber: "Cyberpunk", bitcoin: "Bitcoin", mono: "Mono", bauhaus: "Bauhaus", github: "GitHub", premium: "Premium" })[themeName] || "Dark";
    if (isLight) {
        ToggleBtn.innerHTML = `${sunIconSvg} <span>${label}</span>`;
        ToggleBtn.setAttribute("title", "Switch to dark mode");
    } else {
        ToggleBtn.innerHTML = `${moonIconSvg} <span>${label}</span>`;
        ToggleBtn.setAttribute("title", "Switch to Light Mode");
    }
}

function getCurrentTheme() {
    const t = document.documentElement.getAttribute("data-theme");
    return (t === "linear" || t === "terminal" || t === "dopamine" || t === "amoled" || t === "light" || t === "retro" || t === "organic" || t === "cyber" || t === "bitcoin" || t === "mono" || t === "bauhaus" || t === "github" || t === "premium") ? t : "linear";
}

function isLightTheme() {
    const t = getCurrentTheme();
    if (t === "github") return document.body.classList.contains('github-light');
    return t === "light" || t === "organic" || t === "mono" || t === "bauhaus";
}

function updateThemeMenu(themeName) {
    document.querySelectorAll('.theme-option[data-theme-option]').forEach(el => {
        el.classList.toggle('active', el.getAttribute('data-theme-option') === themeName);
    });
}

// Live theme picker: per-theme palette dots (click commits)
let activeTheme = "linear";

const THEME_PALETTES = {
    linear:   ["#5E6AD2", "#6872D9", "#a78bfa", "#18181b"],
    amoled:   ["#10b981", "#34d399", "#000000", "#27272a"],
    light:    ["#3b82f6", "#2563eb", "#ffffff", "#e2e8f0"],
    terminal: ["#33ff00", "#66ff33", "#0a0a0a", "#1f521f"],
    dopamine: ["#FF3AF2", "#7B2FFF", "#FF6BFF", "#0D0D1A"],
    retro:    ["#0000FF", "#00ff00", "#ff8000", "#C0C0C0"],
    organic:  ["#5D7052", "#6E8260", "#DED8CF", "#FDFCF8"],
    cyber:    ["#00ff88", "#4dffab", "#ff2a6d", "#0a0a0f"],
    bitcoin:  ["#F7931A", "#FFA93F", "#94A3B8", "#1E293B"],
    mono:     ["#000000", "#525252", "#E5E5E5", "#FFFFFF"],
    bauhaus:  ["#D02020", "#1040C0", "#F0C020", "#121212"],
    github:   ["#58a6ff", "#3fb950", "#f0f6fc", "#0d1117"],
    premium:  ["#28E0C7", "#FFFFFF", "#A6A6A6", "#131313"]
};

function renderThemeSwatches() {
    document.querySelectorAll('.theme-chip[data-theme]').forEach(chip => {
        const palette = THEME_PALETTES[chip.getAttribute('data-theme')];
        const swatch = chip.querySelector('.tc-swatch');
        if (!palette || !swatch) return;
        swatch.innerHTML = "";
        palette.forEach(color => {
            const dot = document.createElement('span');
            dot.className = 'tc-dot';
            dot.style.backgroundColor = color;
            swatch.appendChild(dot);
        });
    });
}

function initThemePicker() {
    renderThemeSwatches();
}

function applyChartTheme(themeName) {
    if (!window.Chart) return;
    if (themeName === "light") {
        Chart.defaults.color = '#334155';
        Chart.defaults.borderColor = '#cbd5e1';
    } else if (themeName === "terminal") {
        Chart.defaults.color = '#33ff00';
        Chart.defaults.borderColor = '#1f521f';
    } else if (themeName === "dopamine") {
        Chart.defaults.color = '#FFFFFF';
        Chart.defaults.borderColor = '#7B2FFF';
    } else if (themeName === "retro") {
        Chart.defaults.color = '#00ff00';
        Chart.defaults.borderColor = '#1084d0';
    } else if (themeName === "organic") {
        Chart.defaults.color = '#78786C';
        Chart.defaults.borderColor = '#E6DCCD';
    } else if (themeName === "cyber") {
        Chart.defaults.color = '#00ff88';
        Chart.defaults.borderColor = '#1c1c2e';
    } else if (themeName === "bitcoin") {
        Chart.defaults.color = '#94A3B8';
        Chart.defaults.borderColor = '#1E293B';
    } else if (themeName === "mono") {
        Chart.defaults.color = '#525252';
        Chart.defaults.borderColor = '#E5E5E5';
    } else if (themeName === "bauhaus") {
        Chart.defaults.color = '#121212';
        Chart.defaults.borderColor = '#121212';
    } else if (themeName === "github") {
        if (document.body.classList.contains('github-light')) {
            Chart.defaults.color = '#59636e';
            Chart.defaults.borderColor = '#d0d7de';
        } else {
            Chart.defaults.color = '#8d96a0';
            Chart.defaults.borderColor = '#30363d';
        }
    } else if (themeName === "premium") {
        Chart.defaults.color = '#A6A6A6';
        Chart.defaults.borderColor = 'rgba(255,255,255,0.06)';
    } else {
        Chart.defaults.color = '#a1a1aa';
        Chart.defaults.borderColor = '#18181b';
    }
}

// Apply a theme to the whole app. Valid themes: linear, terminal, dopamine, amoled, light, retro, organic, cyber, bitcoin, mono, bauhaus, github, premium.
function setTheme(themeName, persist) {
    const valid = ["linear", "terminal", "dopamine", "amoled", "light", "retro", "organic", "cyber", "bitcoin", "mono", "bauhaus", "github", "premium"];
    if (!valid.includes(themeName)) themeName = "linear";
    if (persist === undefined) persist = true;

    const changed = themeName !== getCurrentTheme();
    document.documentElement.setAttribute("data-theme", themeName);
    document.body.classList.toggle('light-mode', themeName === "light");
    document.body.classList.remove('github-light');

    if (persist) {
        localStorage.setItem("AppThemeMode", themeName);
        activeTheme = themeName;
    }

    updateThemeToggleButton(themeName);
    updateThemeMenu(themeName);
    applyChartTheme(themeName);

    if (!changed) return;
    if (typeof refreshDashboard === 'function') refreshDashboard();
    if (typeof applyInsightFilter === 'function') applyInsightFilter();
    if (window.activityTrend && typeof window.activityTrend.renderUI === 'function') window.activityTrend.renderUI();
}

// Initialize theme state on page boot
function initTheme() {
    const saved = localStorage.getItem("AppThemeMode");
    let theme = "linear";
    if (saved === "light") theme = "light";
    else if (saved === "dark" || saved === "amoled") theme = "amoled";
    else if (saved === "terminal") theme = "terminal";
    else if (saved === "dopamine") theme = "dopamine";
    else if (saved === "retro") theme = "retro";
    else if (saved === "organic") theme = "organic";
    else if (saved === "cyber") theme = "cyber";
    else if (saved === "bitcoin") theme = "bitcoin";
    else if (saved === "mono") theme = "mono";
    else if (saved === "bauhaus") theme = "bauhaus";
    else if (saved === "github") theme = "github";
    else if (saved === "premium") theme = "premium";
    else if (saved === "linear") theme = "linear";
    setTheme(theme, false);
    activeTheme = theme;
    initSpotlights();
    initThemePicker();
}

function toggleTheme() {
    const cur = getCurrentTheme();
    if (cur === "github") {
        document.body.classList.toggle('github-light');
        updateThemeToggleButton("github");
        applyChartTheme("github");
        if (typeof refreshDashboard === 'function') refreshDashboard();
        if (typeof applyInsightFilter === 'function') applyInsightFilter();
        if (window.activityTrend && typeof window.activityTrend.renderUI === 'function') window.activityTrend.renderUI();
        return;
    }
    setTheme(cur === "light" ? "amoled" : "light");
}

// KPI card mouse-tracking spotlight (dashboard-tuned)
function initSpotlights() {
    document.querySelectorAll('.kpi-card, .chart-card, .table-card').forEach(card => {
        card.addEventListener('pointermove', (e) => {
            const r = card.getBoundingClientRect();
            card.style.setProperty('--mx', `${e.clientX - r.left}px`);
            card.style.setProperty('--my', `${e.clientY - r.top}px`);
        });
    });
}

// Auto-execute theme initialization on script load & DOMContentLoaded so theme persists across ALL pages
if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initTheme);
} else {
    initTheme();
}

// Render Consignees Chart (> 5 m³ volume, ALL consignees included)
function renderCharts() {
    const consigneeChartElem = document.getElementById('consigneeChart');
    if (!consigneeChartElem) return;

    const consigneeVol = {};
    let courtsOriginalKey = null;
    let courtsHasSrWhse = false;

    DataHoarderArray.forEach(item => {
        const consigneeKey = (item.name && item.name.trim() !== "") ? item.name.trim() : "UNASSIGNED";
        if (consigneeKey !== "UNASSIGNED") {
            consigneeVol[consigneeKey] = (consigneeVol[consigneeKey] || 0) + item.vol;

            // Detect COURTS consignee + Tampines North SR/WHSE address (col F)
            if (consigneeKey.toUpperCase().includes("COURTS")) {
                courtsOriginalKey = consigneeKey;
                if ((item.addr || "").toUpperCase().includes("TAMPINES NORTH")) {
                    courtsHasSrWhse = true;
                }
            }
        }
    });

    // Relabel COURTS bar to flag SR/WHSE direct-delivery hub (merged total volume)
    if (courtsHasSrWhse && courtsOriginalKey && consigneeVol[courtsOriginalKey]) {
        consigneeVol["COURTS (SINGAPORE) PTE LTD (SR/WHSE)"] = consigneeVol[courtsOriginalKey];
        delete consigneeVol[courtsOriginalKey];
    }

    if (TopConsigneeChart) TopConsigneeChart.destroy();

    const sortedConsignees = Object.entries(consigneeVol)
        .filter(item => item[1] > 5)
        .sort((a, b) => b[1] - a[1]);

    const isLight = isLightTheme();
    const gridColor = isLight ? '#e2e8f0' : '#18181b';
    const textColor = isLight ? '#334155' : '#a1a1aa';

    const VibrantColors = [
        '#10b981', // Emerald Green
        '#8b5cf6', // Royal Purple
        '#3b82f6', // Electric Blue
        '#f59e0b', // Sunset Amber
        '#ec4899', // Neon Pink
        '#06b6d4', // Cyan Teal
        '#6366f1'  // Indigo
    ];

    const BorderColors = [
        '#34d399',
        '#a78bfa',
        '#60a5fa',
        '#fbbf24',
        '#f472b6',
        '#22d3ee',
        '#818cf8'
    ];

    const barColors = sortedConsignees.map((_, idx) => VibrantColors[idx % VibrantColors.length]);
    const barBorders = sortedConsignees.map((_, idx) => BorderColors[idx % BorderColors.length]);

    const finalVolData = sortedConsignees.map(item => parseFloat(item[1].toFixed(2)));

    TopConsigneeChart = new Chart(consigneeChartElem, {
        type: 'bar',
        data: {
            labels: sortedConsignees.map(item => item[0]),
            datasets: [{
                label: 'Volume (m³)',
                data: sortedConsignees.map(() => 0), // Start from 0 to trigger growth animation
                backgroundColor: barColors,
                borderColor: barBorders,
                borderWidth: 1.5,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            maxBarThickness: 28,
            animation: {
                duration: 1200,
                easing: 'easeOutQuart'
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: isLight ? '#ffffff' : '#09090b',
                    titleColor: isLight ? '#0f172a' : '#ffffff',
                    bodyColor: isLight ? '#059669' : '#34d399',
                    borderColor: isLight ? '#cbd5e1' : '#27272a',
                    borderWidth: 1,
                    padding: 10,
                    callbacks: {
                        label: function(context) {
                            return ' Volume: ' + context.parsed.x + ' m³ (Direct Delivery Eligible)';
                        }
                    }
                }
            },
            scales: {
                x: { 
                    grid: { color: gridColor }, 
                    ticks: { 
                        color: textColor,
                        callback: function(val) { return val + ' m³'; }
                    } 
                },
                y: { grid: { color: gridColor }, ticks: { color: textColor, font: { weight: '600' } } }
            }
        }
    });

    // Trigger smooth bar growth animation extending from 0 to full volume
    requestAnimationFrame(() => {
        requestAnimationFrame(() => {
            TopConsigneeChart.data.datasets[0].data = finalVolData;
            TopConsigneeChart.update();
        });
    });
}

