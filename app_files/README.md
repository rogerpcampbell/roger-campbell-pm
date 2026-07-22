# Roger Campbell Rail / Ponds / Roads Control Center v35

## Executive analytics update

- The Overview starts with four concise management signals for progress, actions, potential exposure, and baseline slip.
- Headline cards open the relevant scope or cost panel; the portfolio action card opens a copyable action list.
- Progress, action pressure, and cost exposure use separate single-axis charts so units and comparisons remain clear.
- A data-freshness strip distinguishes the active weekly report, monthly cost report, and schedule baseline dates.
- The six-week control pressure heatmap and all existing scope detail, action, cost, and milestone sections remain available.
- Weekly controls now include the 2026 Week 27 report with a 04 Jul 2026 cut-off; prior weekly history remains selectable.

Windows launch:

1. Unzip this package.
2. Open the extracted folder.
3. Double-click `RUN_APP_WINDOWS.bat`.

## Main update in v29

This version keeps the app focused on **Rail On Site (ROS)**, **Ponds**, and **Roads** only, with UGP fully removed.

Key v29 changes:

- The title/control box is still visible while scrolling, but it is now sticky inside the app content so it is not clipped by the Streamlit sidebar.
- Header controls change depending on the selected panel:
  - Executive overview: actual week, Back / Next week, and cut-off only.
  - Cost Status: cost month, cost view, Back / Next month, and active report month only.
  - Scope panels: actual week, baseline, Back / Next week, and cut-off only.
- Cost Status Back / Next buttons now move through monthly cost reports rather than weekly reports.
- Dropdown lists were restyled so their selected values and dropdown menus remain visible.
- Cost KPI history is now opened by click instead of hover, making the history easier to review.
- Cost wording was clarified: “Reconciliation” was replaced by “Cost check / Balanced / Review”.
- The duplicate ETC health card was replaced with VOWD / Forecast in Cost Status.
- Executive overview now includes a compact forecast / cost context section showing construction forecast, engineering forecast, Forecast, VOWD, and ETC by scope.
- Spacing between sections was tightened and made more content-driven.
- The dedicated **Cost Status** panel remains available for deeper cost control, history, KPIs, waterfall/candle bars, package drivers, and exposure items.

## Existing capabilities retained

- Executive overview for Rail On Site (ROS), Ponds, and Roads only.
- Detailed panels for Rail On Site (ROS), Ponds, and Roads.
- Milestone timelines starting at the active week and ordered by forecast/control date.
- Schedule baseline control per scope, collapsed by default and labelled with the active baseline.
- Upload controls for weekly reports, monthly cost reports, and schedule baseline PDFs.
- Cost Status can visualize All scopes, Rail On Site (ROS), Ponds, or Roads.
- Table copy behavior tries a formatted HTML table first, with Excel-ready text fallback if the browser blocks rich clipboard access.

The source files used for traceability are in `sources/`.
