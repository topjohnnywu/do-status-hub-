# DO Status Hub

Multi-page analytics dashboard for Delivery Order (DO) status tracking, truck planning, batch analytics and shipping insight. Pure **HTML5 + CSS3 + vanilla JavaScript (ES6+)** — no framework, no build step. Excel files are parsed directly in the browser with SheetJS, and charts are rendered with Chart.js.

## ✨ Recent Updates & Enhancements

- **📊 Dynamic Filtered KPI Cards (DO Summary Generator)**:
  - KPI cards (Total DOs, Volume, Quantity, SKU) now recalculate in real-time when search filters or text queries are applied to DO records.
  - Automatic dynamic indicator labels switch between *Parsed Orders* and *Filtered Orders* during active search sessions.
- **💻 Laptop-Friendly Responsive 2-Column Grid**:
  - Re-architected the DO Summary Generator header layout into a responsive 2-column grid (`1fr 340px`/`380px`), ensuring high readability and minimal vertical scrolling on laptops and wide screens.
- **⚡ Quick Remarks & Presets Side Panel**:
  - Added a collapsible sticky side panel for 1-click bulk remark application (`SELF COLLECT`, `HOLD`, `LOCAL DELIVERY`, `DIRECT DELIVERY`, `URGENT`, `CANCELLED`, `CLEAR REMARK`).
  - Integrated custom bulk remark input for fast multi-row assignment and quick manual DO row additions.
- **🔄 Supplementary Missing Info Updater**:
  - Added secondary CSV/Excel file import ("Update Missing Info") to seamlessly patch missing customer addresses, routes, or item details across existing DO batches without overwriting manual edits.

## Features

- **In-browser Excel import** — upload DO Summary, Batch Picking (`.xlsm`) and Shipping Insight files (`.xlsx`, `.xls`, `.csv`) via drag-and-drop file pickers; everything is parsed client-side with SheetJS.
- **8 dedicated modules** — summary analytics, shipping insight, DO details inspector, truck planning, batch analytics, summary generator, challenger list and activity trend.
- **13 switchable themes** — from a Linear-style dark UI to Terminal, Cyberpunk, Bitcoin, Bauhaus, GitHub and a Premium enterprise look. The active theme persists across sessions (`localStorage` key `AppThemeMode`, default: Linear).
- **Theme-aware charts** — every chart palette, axis, grid and legend re-skins live when the theme changes.
- **Interactive tables** — sortable/filterable grid with in-cell editing, inline remark/comment inputs, missing-value highlighting and search.
- **No backend required** — the whole app runs from static files and can be opened straight from disk or served by any static server.

## Pages

| Module | File | Purpose |
|---|---|---|
| Summary Analytics | `index.html` | Dashboard home; uploads, KPI overview and top-consignee charts |
| Shipping Insight | `shipping_insight.html` | Shipping-volume analytics and insights |
| DO Details Inspector | `do_details.html` | Drill into individual delivery-order records |
| Truck Planning | `truck_planning.html` | Daily truck/summary manifest planning with batch tab and in-table editing |
| Batch Analytics | `batch_analytics.html` | Batch-picking analytics and charts |
| DO Summary Generator | `do_summary_generator.html` | Compile, filter, annotate, and export DO summary data |
| Challenger List | `challenger_list.html` | Track and filter challenger deliveries |
| DO Activity Trend | `do_activity_trend.html` | Activity-over-time trend charts |

All pages share the same layout (sidebar navigation, theme picker), one stylesheet and one theme engine, so switching a theme restyles the entire app instantly.

## Theming

Themes are defined as CSS custom-property blocks in `css/styles.css` (`[data-theme="..."]`) plus chart palettes in `js/charts.js`. The theme engine (`setTheme`, `initTheme`, `toggleTheme`, `applyChartTheme`) lives in `js/charts.js` and keeps the chart rendering in sync with the selected theme.

| Theme | Style |
|---|---|
| Linear | Indigo-on-charcoal dark UI |
| Terminal | Monospace green-on-black |
| Dopamine | Pink/magenta gradient dark |
| Retro | Warm orange/pink gradient |
| AMOLED | True-black with high contrast |
| Light | White + slate, classic light mode |
| Organic | Green/natural light palette |
| Cyberpunk | Neon cyan/magenta on deep purple |
| Bitcoin | Gold/black with orange accents |
| Mono | Minimalist monochrome |
| Bauhaus | Primary-color light design |
| GitHub | GitHub Dark with a Light variant (auto via `data-color-mode`) |
| Premium | Ultra-dark enterprise AI-dashboard look with a teal accent (`#28E0C7`) |

## Tech Stack

- **HTML5 / CSS3** — semantic markup, CSS custom properties for theming
- **Vanilla JavaScript (ES6+)** — modular per-page scripts
- **Chart.js** — charting (loaded from CDN)
- **SheetJS (xlsx)** — client-side Excel parsing (loaded from CDN)
- **VBA** — optional legacy macros (`*.bas`) for Excel-side automation (bulk import, challenger formatting, email)

## Getting Started

There is no build step or install. Either:

1. Open `index.html` directly in a modern browser, or
2. Serve the folder with any static file server, e.g.

```bash
python -m http.server 8000
```

Then open `http://localhost:8000` and upload a DO Summary file (`.xlsx`) to start.

## Pushing Changes to GitHub

To push your latest changes to your remote GitHub repository, run the following standard git commands in your terminal:

```bash
# 1. Stage all updated files
git add .

# 2. Commit your changes with a descriptive message
git commit -m "feat: add dynamic KPI filtering, quick remarks panel, and laptop-friendly layout to DO summary generator"

# 3. Push to your main branch (e.g., main or master)
git push origin main
```

## Project Structure

```
├── index.html                  # Dashboard home
├── shipping_insight.html       # Shipping insight module
├── do_details.html             # DO details inspector
├── truck_planning.html         # Truck planning + daily manifest
├── batch_analytics.html        # Batch analytics
├── do_summary_generator.html   # DO summary generator
├── challenger_list.html        # Challenger list
├── do_activity_trend.html      # Activity trend
├── css/
│   └── styles.css              # Single shared stylesheet with 13 theme blocks
├── js/
│   ├── app.js                  # App shell / shared helpers
│   ├── charts.js               # Theme engine + chart palettes
│   ├── parsers.js              # Excel parsing
│   ├── *.js                    # Per-page module scripts
│   └── xlsx.bundle.js          # SheetJS bundle
└── *.bas                       # Optional Excel/VBA helper macros
```

## Browser Support

Targets current evergreen browsers (Chrome, Edge, Firefox, Safari). File handling uses the File API; charts require canvas support.
