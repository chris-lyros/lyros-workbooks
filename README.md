# Lyros Workbooks

Free Excel finance workbooks for Australian businesses on Xero, with a plain-English finder.

**Live:** https://workbooks.lyros.com.au

## What this is

Describe the finance workbook you need in plain English and the finder matches it to a free, ready-to-use XLSX from the Lyros library. Every workbook is built in-house by Lyros Accounting around Australian Xero data and ships with synthetic sample figures, so it is safe to download and adapt.

## Repository structure

| Path | Purpose |
| --- | --- |
| `index.html`, `styles.css`, `tokens.css`, `app.js` | The static site, served as-is |
| `assets/`, `fonts/` | Branding and typography |
| `library/` | The published `.xlsx` workbooks served to visitors |
| `build/` | Python (openpyxl) generators that produce everything in `library/` |
| `wrangler.jsonc` | Cloudflare Workers static-asset configuration |

## Deployment

The site runs on Cloudflare Workers (static assets) and deploys automatically on every push to `main`. No build step runs on the host: the workbooks in `library/` are generated locally and committed, and Cloudflare serves the repository root as-is.

## Regenerating the workbooks

The `.xlsx` files are produced locally with Python and openpyxl:

```bash
cd build
python _run_all.py
```

This rebuilds every workbook in `library/` and the library index. Commit the updated files and push, and Cloudflare redeploys.

## Disclaimer

The workbooks are general tools and do not replace tailored accounting, tax or financial advice. They contain synthetic sample data. Connecting your own live Xero data to a workbook is a Lyros advisory service: see https://www.lyros.com.au/contact.

(c) Lyros Pty Ltd (ABN 46 689 015 165), trading as Lyros Accounting.
