# Incident Tracker Dashboard

This is a simple static dashboard that lets you upload an Excel file (Incident Tracker) and view multiple charts to help managers understand site safety.

How it works

- Open index.html in a browser (or host on GitHub Pages).
- Click "Choose Excel file" and pick your Incident Tracker .xlsx file.
- The site parses the first sheet and expects columns with these headers (case-insensitive):
  SR, Site, Sources, Shift, Year, Date, Incident Category, Criticality, Type of Incident, Event Title, GEHSMS Standard, Detailed Observation, What is the Action?, Priority, Owner, Status, New Timeline, Department, Area, Contractor Name, Responsible
- It then shows a table and multiple charts (Incidents by Site, by Year, by Incident Category, by Criticality, by Status).

Files added

- index.html — the dashboard UI
- app.js — parsing, aggregation, chart rendering logic
- styles.css — small styles
- README.md — usage and deployment notes

Feel free to ask for more charts or back-end integration (for uploads, persistence, authentication).