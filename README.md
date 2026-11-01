# OrphLux Data Catalog

Public catalog of OrphLux datasets and data projects, live at
https://orphlux-data-catalog.vercel.app

## How it works

- **Content source:** the public Google Sheet "OrphLux Dataset List" (Drive file
  `1cHy8M1_U9icoOho4fCpz79fWqDS_madQ`, shared as *Anyone with the link – Viewer*).
  All content edits happen in that sheet; no code change or redeploy is needed.
- **`api/catalog.py`:** Vercel Python function. It downloads the sheet as .xlsx,
  parses it, and returns JSON at `/api/catalog`. Vercel's CDN caches the response
  for 1 hour (`s-maxage=3600`, `stale-while-revalidate=86400`), so sheet edits
  appear on the site within about an hour. If the download fails or the result
  looks wrong (<200 rows or an empty tab), it returns 503.
- **`index.html`:** single-file front end (no build step). It loads `/api/catalog`,
  and on failure falls back to the bundled snapshot in `d/p0.bin … p9.bin`
  (gzip JSON split into 10 parts). `window.__catalogSource` is `live` or `snapshot`.
- **Layout:** three tabs, each with sidebar types (mapping lives in `classify()` in `index.html`):
  - **For Generative AI** — Coding Data · STEM Data · Expert Data · Benchmark Data · Multimodal Data (Image / Video / Audio / Multimodal Data Annotation) · Others
  - **For Physical AI** — Ego Data · UMI Data · Teleoperation Data · Others
  - **For World Model**
  Text datasets are split into Coding / STEM / Expert / Others by keywords in their names
  (`RX_EXPERT`, `RX_STEM`); anything unmatched lands in Others.

- **Descriptions:** a card's description comes from the sheet's description cell. A cell that only
  contains an attached file name (e.g. `Spec.docx`) is ignored. When the cell is empty, the page uses a
  summary written from that entry's overview document (the `DESC` table near the top of the script in
  `index.html`); otherwise the card shows no description. Typing a description into the sheet always wins.

## Files

| Path | Purpose |
|---|---|
| `index.html` | The whole site (HTML/CSS/JS) |
| `api/catalog.py` | Live sheet → JSON endpoint |
| `requirements.txt` | `openpyxl` for the endpoint |
| `vercel.json` | Function config (30 s max duration) |
| `d/` | Fallback snapshot of the catalog + `manifest.json` |
| `scripts/build_catalog.py` | Rebuilds the `d/` snapshot from a downloaded .xlsx (same parsing rules as the endpoint) |

## Refreshing the fallback snapshot (optional)

```bash
pip install openpyxl
python scripts/build_catalog.py "OrphLux Dataset List.xlsx" d/   # args: <source.xlsx> <output dir>
```

## Deploy

Vercel project `orphlux-data-catalog`. Static files + one Python function; no
framework or build command.
