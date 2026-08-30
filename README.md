# EHS Tracker Dashboards

This static, client-side dashboard turns EHS Excel workbooks into charts and searchable data previews. Each supported tracker has its own upload control and dashboard section.

How it works

- Open index.html in a browser (or host on GitHub Pages).
- Choose the upload card for the relevant tracker and select an `.xlsx` or `.xls` file.
- The first worksheet is parsed in the browser; files are not sent to a server.
- Upload any combination of Incident, BBS, MOC, GEHMS, EHS Observations, and PTW trackers.
- Each workbook appears in a separate section with four tracker-specific charts and a collapsible table preview.
- Header matching is case-insensitive and ignores repeated whitespace, periods, and underscores.

Files added

- index.html — the dashboard UI
- app.js — parsing, aggregation, chart rendering logic
- styles.css — small styles
- README.md — usage and deployment notes

Feel free to ask for more charts or back-end integration (for uploads, persistence, authentication).
