# Work samples

Writing, translation and spreadsheet-automation samples. All stories, brands and data are original or fictional.

**Live gallery:** https://rafa0071.github.io/work-samples/

## Creative writing
- `writing/pip-shares-a-little.html`: picture-book manuscript (ages 3–5) with spreads and read-aloud notes
- `writing/blackout-on-line-9.html`: detective novel, chapter one
- `writing/la-luna-que-no-queria-dormir.html`: Spanish bedtime story
- `writing/the-lighthouse-ledger.html`: interactive murder-mystery case file

## Translation
- `translation/website-en-es.html`: restaurant website, EN → ES
- `translation/app-strings-json.html`: i18n JSON, EN → ES / PT-BR
- `translation/subtitles-srt.html`: video subtitles, EN → ES
- `translation/product-sheet-es-en.html`: coffee product sheet, ES → EN (US)

## Excel & data automation (Node.js + ExcelJS)
Each folder contains the script, its input and the real output it produced.

| Folder | What it does |
| --- | --- |
| `automation/sales-report` | Merges monthly CSVs into a formatted Excel report with SUM formulas |
| `automation/clean-customers` | Normalizes and deduplicates a messy customer list, with an Issues sheet |
| `automation/invoices` | Turns an Excel billing sheet into one PDF invoice per client |
| `automation/scrape-books` | Scrapes a public practice catalogue into Excel |

```bash
npm install exceljs pdfkit cheerio
node automation/make-sample-data.js
node automation/sales-report/merge-sales.js automation/sales-report/data report.xlsx
```

## License
MIT
