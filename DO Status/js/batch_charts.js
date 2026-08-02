// Batch Analytics Charts - Chart.js Rendering Functions

let chartInstances = {};

// Accessibility: disable animations for users who prefer reduced motion
function prefersReducedMotion() {
    return window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

// Config-driven metric system for dynamic charts
const BATCH_METRICS = {
    volume:  { label: 'Volume (m³)', field: 'totalVolume',  color: '#ef4444', isVolume: true  },
    do:      { label: 'DO',          field: 'totalDO',      color: '#3b82f6', isVolume: false },
    sku:     { label: 'SKU',         field: 'totalSKU',     color: '#f59e0b', isVolume: false },
    trucks:  { label: 'Trucks',      field: 'totalTrucks',  color: '#10b981', isVolume: false }
};

function renderBatchCharts() {
    // Destroy existing charts
    Object.values(chartInstances).forEach(chart => chart?.destroy());
    chartInstances = {};
    
    const batches = batchManager.getFilteredBatches();
    
    if (batches.length === 0) {
        console.log('No batch data available for charts');
        return;
    }
    
    const isDark = !isLightTheme();
    const textColor = isDark ? '#a1a1aa' : '#52525b';
    const gridColor = isDark ? '#27272a' : '#e4e4e7';
    
    renderDODistributionChart(batches, textColor, gridColor);
    renderVolumeTrendChart(batches, textColor, gridColor);
    renderTopBatchesChart(batches, textColor, gridColor);
    renderTrucksHorizontalChart(batches, textColor, gridColor);
}

// Chart 1: DO Type Distribution
function renderDODistributionChart(batches, textColor, gridColor) {
    const ctx = document.getElementById('doDistributionChart');
    if (!ctx) return;
    
    const totalSmall = batches.reduce((sum, b) => sum + b.smallDOCount, 0);
    const totalBig = batches.reduce((sum, b) => sum + b.bigDOCount, 0);
    const hasDataLabels = typeof ChartDataLabels !== 'undefined';
    
    chartInstances.doDistribution = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Small DO', 'Big DO'],
            datasets: [{
                data: [totalSmall, totalBig],
                backgroundColor: ['#f59e0b', '#3b82f6'],
                borderColor: isDarkMode() ? '#18181b' : '#fff',
                borderWidth: 2,
                hoverOffset: 10
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: 12 },
            animation: prefersReducedMotion() ? false : {
                animateRotate: true,
                animateScale: true,
                duration: 1000,
                easing: 'easeOutQuart'
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'bottom',
                    labels: { color: textColor, font: { size: 13 }, padding: 14 }
                },
                tooltip: {
                    backgroundColor: isDarkMode() ? '#18181b' : '#fff',
                    titleColor: textColor,
                    bodyColor: textColor,
                    borderColor: gridColor,
                    borderWidth: 1,
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.parsed / total) * 100).toFixed(1);
                            return `${context.label}: ${context.parsed.toLocaleString()} (${percentage}%)`;
                        }
                    }
                },
                datalabels: {
                    display: function(context) {
                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                        return total > 0 && (context.dataset.data[context.dataIndex] / total) >= 0.05;
                    },
                    formatter: function(value, context) {
                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                        return ((value / total) * 100).toFixed(1) + '%';
                    },
                    color: '#ffffff',
                    font: { size: 15, weight: 'bold' },
                    textAlign: 'center',
                    textStrokeColor: 'rgba(0, 0, 0, 0.6)',
                    textStrokeWidth: 4
                }
            }
        },
        plugins: hasDataLabels ? [ChartDataLabels] : []
    });
}

// Chart 2: Performance Trend (metric selectable via dropdown)
function renderVolumeTrendChart(batches, textColor, gridColor) {
    const ctx = document.getElementById('volumeTrendChart');
    if (!ctx) return;
    
    // Read selected metric from dropdown, default to 'volume'
    const metricKey = document.getElementById('trendMetricFilter')?.value || 'volume';
    const metric = BATCH_METRICS[metricKey];
    if (!metric) return;
    
    const sorted = [...batches].sort((a, b) => new Date(a.batchDate) - new Date(b.batchDate));
    
    // Destroy existing chart if it exists
    if (chartInstances.volumeTrend) {
        chartInstances.volumeTrend.destroy();
    }
    
    chartInstances.volumeTrend = new Chart(ctx, {
        type: 'line',
        data: {
            labels: sorted.map(b => new Date(b.batchDate).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' })),
            datasets: [{
                label: metric.label,
                data: sorted.map(b => b[metric.field]),
                borderColor: metric.color,
                backgroundColor: metric.color + '1A',
                borderWidth: 2,
                tension: 0.4,
                fill: true,
                pointRadius: 4,
                pointBackgroundColor: metric.color
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animations: prefersReducedMotion() ? {} : {
                y: {
                    duration: 1200,
                    easing: 'easeOutQuart',
                    from: (ctx) => {
                        if (ctx.type === 'data' && ctx.chart && ctx.chart.scales.y) {
                            return ctx.chart.scales.y.getPixelForValue(0);
                        }
                    },
                    delay: (ctx) => {
                        if (ctx.type === 'data' && ctx.dataIndex !== undefined) {
                            return ctx.dataIndex * 60;
                        }
                        return 0;
                    }
                }
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: { color: textColor, font: { size: 11 } }
                },
                tooltip: {
                    backgroundColor: isDarkMode() ? '#18181b' : '#fff',
                    titleColor: textColor,
                    bodyColor: textColor,
                    borderColor: gridColor,
                    borderWidth: 1,
                    callbacks: {
                        label: function(context) {
                            const val = context.parsed.y;
                            if (metric.isVolume) {
                                return `${metric.label}: ${val.toFixed(2)} m³`;
                            }
                            return `${metric.label}: ${val.toLocaleString()}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: { color: textColor, font: { size: 10 } },
                    grid: { display: false }
                },
                y: {
                    ticks: { 
                        color: textColor, 
                        font: { size: 10 },
                        callback: function(value) {
                            if (metric.isVolume) {
                                return value.toFixed(1) + ' m³';
                            }
                            return value.toLocaleString();
                        }
                    },
                    grid: { color: gridColor }
                }
            }
        }
    });
}

// Chart 3: Top Batches by [metric] (metric selectable via dropdown)
function renderTopBatchesChart(batches, textColor, gridColor) {
    const ctx = document.getElementById('topBatchesChart');
    if (!ctx) return;
    
    // Read selected metric from dropdown, default to 'sku'
    const metricKey = document.getElementById('topBatchesMetricFilter')?.value || 'sku';
    const metric = BATCH_METRICS[metricKey];
    if (!metric) return;
    
    // Sort by chosen metric, descending
    const topBatches = [...batches]
        .sort((a, b) => b[metric.field] - a[metric.field])
        .slice(0, 10);
    
    // Destroy existing chart if it exists
    if (chartInstances.topBatches) {
        chartInstances.topBatches.destroy();
    }
    
    // Multi-color palette for vibrant, distinct bar graph visualization
    const barColors = [
        '#3b82f6', // Electric Blue
        '#10b981', // Emerald Green
        '#8b5cf6', // Royal Purple
        '#f59e0b', // Amber Yellow
        '#ec4899', // Neon Pink
        '#06b6d4', // Cyan Teal
        '#f97316', // Bright Orange
        '#6366f1', // Indigo
        '#14b8a6', // Teal
        '#e11d48'  // Rose
    ];

    chartInstances.topBatches = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: topBatches.map(b => b.batchName),
            datasets: [{
                label: metric.label,
                data: topBatches.map(b => b[metric.field]),
                backgroundColor: topBatches.map((_, idx) => barColors[idx % barColors.length]),
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'x',
            animation: prefersReducedMotion() ? false : {
                duration: 900,
                easing: 'easeOutQuart',
                delay: (context) => {
                    // Stagger each bar's growth
                    if (context.type === 'data' && context.mode === 'default' && context.dataIndex !== undefined) {
                        return context.dataIndex * 80;
                    }
                    return 0;
                }
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: { color: textColor, font: { size: 11 } }
                },
                tooltip: {
                    backgroundColor: isDarkMode() ? '#18181b' : '#fff',
                    titleColor: textColor,
                    bodyColor: textColor,
                    borderColor: gridColor,
                    borderWidth: 1,
                    callbacks: {
                        label: function(context) {
                            const val = context.parsed.y;
                            if (metric.isVolume) {
                                return `${metric.label}: ${val.toFixed(2)} m³`;
                            }
                            return `${metric.label}: ${val.toLocaleString()}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: { color: textColor, font: { size: 10 } },
                    grid: { display: false }
                },
                y: {
                    ticks: { 
                        color: textColor, 
                        font: { size: 10 },
                        callback: function(value) {
                            if (metric.isVolume) {
                                return value.toFixed(1) + ' m³';
                            }
                            return value.toLocaleString();
                        }
                    },
                    grid: { color: gridColor }
                }
            }
        }
    });
}

// Public function: update Performance Trend chart when dropdown changes
function updateTrendChart() {
    const batches = batchManager.getFilteredBatches();
    if (batches.length === 0) return;
    
    const isDark = !isLightTheme();
    const textColor = isDark ? '#a1a1aa' : '#52525b';
    const gridColor = isDark ? '#27272a' : '#e4e4e7';
    
    renderVolumeTrendChart(batches, textColor, gridColor);
}

// Public function: update Top Batches chart when dropdown changes
function updateTopBatchesChart() {
    const batches = batchManager.getFilteredBatches();
    if (batches.length === 0) return;
    
    const isDark = !isLightTheme();
    const textColor = isDark ? '#a1a1aa' : '#52525b';
    const gridColor = isDark ? '#27272a' : '#e4e4e7';
    
    renderTopBatchesChart(batches, textColor, gridColor);
    
    // Update title dynamically
    const metricKey = document.getElementById('topBatchesMetricFilter')?.value || 'sku';
    const metric = BATCH_METRICS[metricKey];
    const titleEl = document.getElementById('topBatchesChartTitle');
    if (titleEl && metric) {
        titleEl.textContent = `Top Batches by ${metric.label.replace(' (m³)', '')}`;
    }
}

// Helper: Generate dynamic, vibrant randomized bar colors for each bar
function getRandomizedBarColors(count) {
    const hueOffset = Math.floor(Math.random() * 360);
    const colors = [];
    for (let i = 0; i < count; i++) {
        // Golden ratio angle distribution + random jitter for distinct vibrant colors
        const hue = (hueOffset + (i * 137.508) + (Math.random() * 20 - 10)) % 360;
        const saturation = 75 + Math.floor(Math.random() * 20); // 75-95%
        const lightness = 52 + Math.floor(Math.random() * 16);  // 52-68%
        colors.push(`hsla(${Math.floor(hue)}, ${saturation}%, ${lightness}%, 0.85)`);
    }
    return colors;
}

// Chart 4: Horizontal Bar Chart for Trucks Analytics with Randomized Bar Colors
function renderTrucksHorizontalChart(batches, textColor, gridColor) {
    const ctx = document.getElementById('trucksHorizontalChart');
    if (!ctx) return;
    
    const metricKey = document.getElementById('truckChartMetricFilter')?.value || 'trucks';
    
    let chartBatches = [...batches];
    let label = 'Total Trucks';
    let getValue = b => b.totalTrucks;
    let formatVal = v => `${v.toLocaleString()} trucks`;

    if (metricKey === 'do') {
        label = 'Total DO';
        getValue = b => b.totalDO;
        formatVal = v => `${v.toLocaleString()} DOs`;
    } else if (metricKey === 'volume') {
        label = 'Total Volume';
        getValue = b => b.totalVolume;
        formatVal = v => `${v.toFixed(2)} m³`;
    } else if (metricKey === 'avgDO') {
        label = 'Avg DO per Truck';
        getValue = b => b.avgDOPerTruck;
        formatVal = v => `${v.toFixed(2)} DO/truck`;
    }

    // Sort batches by selected metric descending
    chartBatches.sort((a, b) => getValue(b) - getValue(a));
    
    // Limit to top 15 batches for optimal visual density
    if (chartBatches.length > 15) {
        chartBatches = chartBatches.slice(0, 15);
    }

    if (chartInstances.trucksHorizontal) {
        chartInstances.trucksHorizontal.destroy();
    }

    const barColors = getRandomizedBarColors(chartBatches.length);

    chartInstances.trucksHorizontal = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: chartBatches.map(b => b.batchName),
            datasets: [{
                label: label,
                data: chartBatches.map(b => getValue(b)),
                backgroundColor: barColors,
                borderColor: barColors.map(c => c.replace('0.85', '1')),
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y', // Horizontal Bar Chart
            responsive: true,
            maintainAspectRatio: false,
            animation: prefersReducedMotion() ? false : {
                duration: 900,
                easing: 'easeOutQuart',
                delay: (context) => {
                    if (context.type === 'data' && context.mode === 'default' && context.dataIndex !== undefined) {
                        return context.dataIndex * 40;
                    }
                    return 0;
                }
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: { color: textColor, font: { size: 11 } }
                },
                tooltip: {
                    backgroundColor: isDarkMode() ? '#18181b' : '#fff',
                    titleColor: textColor,
                    bodyColor: textColor,
                    borderColor: gridColor,
                    borderWidth: 1,
                    callbacks: {
                        label: function(context) {
                            return `${label}: ${formatVal(context.parsed.x)}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: { 
                        color: textColor, 
                        font: { size: 10 },
                        callback: function(val) {
                            if (metricKey === 'volume') return val.toFixed(1) + ' m³';
                            return val.toLocaleString();
                        }
                    },
                    grid: { color: gridColor }
                },
                y: {
                    ticks: { color: textColor, font: { size: 11, weight: '500' } },
                    grid: { display: false }
                }
            }
        }
    });
}

// Public function: update Trucks Horizontal chart when dropdown changes
function updateTrucksChart() {
    const batches = batchManager.getFilteredBatches();
    if (batches.length === 0) return;
    
    const isDark = !isLightTheme();
    const textColor = isDark ? '#a1a1aa' : '#52525b';
    const gridColor = isDark ? '#27272a' : '#e4e4e7';
    
    renderTrucksHorizontalChart(batches, textColor, gridColor);
}

function isDarkMode() {
    return !isLightTheme();
}
