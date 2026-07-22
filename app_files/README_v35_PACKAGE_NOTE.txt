Roger Campbell Rail / Ponds / Roads Control Center v35

Latest UX correction - cost charts and floating panel title
- Fixed the first panel opening too low on the page; the floating title now starts at the top.
- The floating title now shows the active panel name instead of "Roger Campbell Scopes Control Center".
- Sidebar branding now reads "Roger Campbell PM".
- Cost candle bars are restored as true waterfall/candle bars and end with negative VOWD and derived ETC.
- VOWD is plotted as a negative movement; ETC is calculated as Forecast + VOWD movement.
- Cost click details now open compact bar charts with year-month on X and amount or percent on Y.
- Scope construction progress charts now show ETC inside the chart.

Latest data inclusion
- Added BOD-SIN-OVE-PMG-REPGE-11094 Weekly Progress Report BOP-IDOM-Week 25 2026.pdf.
- Added BOD-SIN-OVE-PMG-REPGE-11095 Weekly Progress Report BOP-IDOM-Week 26 2026.pdf.
- The app now starts from latest weekly data: 2026 W26.
- Streamlit now invalidates the weekly-data cache whenever the published bundle changes.
- Week 26 added 1 HSE row, 1 schedule row, 12 waypoints, 1 risk summary row, 6 risk rows, 127 watchlist/action rows, and 3 scope engineering rows.
- Roger can search the complete extracted Week 26 report text as well as the structured dashboard data.
- The source PDF is included in sources/ for traceability.

Latest dropdown correction
- Fixed Streamlit dropdown menus appearing behind the sticky header/cards.
- Raised both the BaseWeb dropdown menu and its Streamlit portal wrapper to the top layer.
- Verified Cost Status Month, View, and Baseline menus are visible and selectable.
- Verified Executive Week, Month, and Baseline menus are visible; the long Week list scrolls to 2026 W26.

Live checks completed
- Executive opened at the top with the active panel title visible.
- Sidebar shows Roger Campbell PM.
- Cost waterfall check confirms Budget, CO, Forecast, Exposure, Potential FC, negative VOWD, and ETC.
- ETC check confirms ETC = Forecast + VOWD movement.
- Weekly bundle check confirms latest period is 2026 W26 with 22 source reports.
- Dropdown check confirms Month can change 2026-May -> 2026-April -> 2026-May.
- Dropdown check confirms View can change All scopes -> Ponds -> All scopes.
- Dropdown check confirms Week can be restored to 2026 W26 from the long list.
- Cost click history check confirms SVG bar charts for amount and percent views.
- App syntax check completed successfully.

Run this corrected package, not the old nested Downloads v29 folder.
