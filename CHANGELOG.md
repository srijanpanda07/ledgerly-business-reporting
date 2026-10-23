# Changelog

## Interface refresh

- Replaced the neon gradients and decorative motion with a calm sage-and-ink visual system.
- Simplified typography, cards, navigation, buttons, and chart colors for a more practical business-reporting feel.
- Updated the app description and deployment instructions for Streamlit Community Cloud.

## Version 2.2 — Vibrant high-contrast UI

- Replaced the low-contrast pale sidebar with a saturated purple gradient, bright lime labels, bold navigation and strong magenta accents.
- Added project-wide dark theme configuration so built-in widgets and table surfaces use the same high-contrast palette.
- Applied bold Manrope / Space Grotesk typography with vivid cyan, magenta, lime, orange and violet accents instead of grey-on-white styling.
- Updated Plotly title, axis, tick, hover and legend colours for dark surfaces and made category/product bars multicolour.
- Added additional entrance motion, interactive hover transitions, button feedback, a subtle animated hero sheen and reduced-motion support.

# Changelog

## Version 2.1 — Readable charts and smooth interactions

- Improved chart text contrast, axis font sizes, hover labels and automatic margins.
- Reworked category and product charts as horizontal bars with value labels, extra right-side breathing room, and full names in hover details to reduce clipped labels.
- Added a selectable Area / Line / Bars view for the monthly sales trend.
- Added smooth Plotly transitions, chart toolbar zoom/scroll interaction, and stable chart keys.
- Added subtle interface motion for cards, charts, metrics, controls and buttons, with a reduced-motion accessibility fallback.
- Added guidance captions for chart interactions and safer axis bounds for negative sales/refund values.

## Version 2.0 — Portfolio upgrade

- Reworked the app into Overview, Explore data, Data quality, and Export center views.
- Added responsive styling, KPI cards, interactive charts, rule-based highlights, shared date/category filters, and table search/pagination.
- Added column mapping for direct revenue or quantity × unit price, plus day-first date parsing option.
- Added a multi-sheet Excel report and a printable PDF summary.
- Split data-processing and PDF-export logic into separate modules.
- Added tests for parsing, duplicate handling, revenue calculation, day-first dates, and Excel/PDF generation.
- Documented the limitations around date parsing, accounting interpretation, and confidential data.
