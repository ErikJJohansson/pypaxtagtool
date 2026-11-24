# Vendored AG Grid Community

AG Grid Community **32.3.0** (MIT licence), vendored locally so the webapp
works on a plant laptop with no internet access. No CDN links at runtime.

Files (downloaded 2026-11-24 from jsDelivr):

- `ag-grid-community.min.js` — UMD bundle, exposes the `agGrid` global
  <https://cdn.jsdelivr.net/npm/ag-grid-community@32.3.0/dist/ag-grid-community.min.js>
- `ag-grid.css` — core styles
  <https://cdn.jsdelivr.net/npm/ag-grid-community@32.3.0/styles/ag-grid.css>
- `ag-theme-quartz.css` — Quartz theme styles
  <https://cdn.jsdelivr.net/npm/ag-grid-community@32.3.0/styles/ag-theme-quartz.css>

To upgrade, download the same three files for the new version from
`https://cdn.jsdelivr.net/npm/ag-grid-community@<version>/...` and update
the version number in this file and in `tests/test_frontend.py`.
