// Builds the HTML showcase page for each automation sample from its real output files.
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const esc = (s) => String(s ?? '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
const fmt = (v) => (v && typeof v === 'object' && 'result' in v ? v.result : v && typeof v === 'object' && 'formula' in v ? '=' + v.formula : v);
const num = (v) => (typeof v === 'number' ? v.toLocaleString('en-US', { maximumFractionDigits: 2 }) : esc(v));

async function sheetTable(file, name, maxRows = 12, cols) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = name ? wb.getWorksheet(name) : wb.worksheets[0];
  const rows = [];
  ws.eachRow((r, i) => { if (i <= maxRows + 1) rows.push(r.values.slice(1).map(fmt)); });
  const pickCols = (r) => (cols ? cols.map((c) => r[c]) : r);
  const [head, ...body] = rows;
  return `<div class="xl"><div class="xl-bar">${esc(path.basename(file))} · sheet "${esc(ws.name)}" · ${ws.rowCount - 1} rows</div><table><tr>${pickCols(head).map((h) => `<th>${esc(h)}</th>`).join('')}</tr>${body.map((r) => `<tr>${pickCols(r).map((c) => `<td>${num(c)}</td>`).join('')}</tr>`).join('')}</table></div>`;
}

function page({ slug, title, kicker, meta, intro, before, after, code, notes }) {
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)} (automation sample)</title>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,700&family=Source+Sans+3:wght@400;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="../../style.css">
<style>
  :root { --accent: #1f7a4d; --bg: #e4ece7; --paper: #fdfefd; }
  .xl { border: 1px solid #cfe0d6; border-radius: 12px; overflow: hidden; margin: 10px 0 18px; background: #fff; }
  .xl-bar { background: #1f7a4d; color: #fff; font-size: 13px; font-weight: 600; padding: 8px 14px; }
  .xl table { font-size: 13px; } .xl th { background: #eef5f1; color: #1d2b24; } .xl td, .xl th { border-right: 1px solid #e5eee9; white-space: nowrap; }
  .xl { overflow-x: auto; }
  .flow { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin: 6px 0 20px; font-weight: 600; }
  .flow span { background: #fff; border: 1px solid #cfe0d6; border-radius: 99px; padding: 6px 14px; font-size: 14px; }
  pre { max-height: 420px; }
</style>
</head>
<body>
<div class="top"><a href="../../">← All samples</a><span>Automation sample · script and real output</span></div>
<article class="doc">
  <div class="kicker">${esc(kicker)}</div>
  <h1>${esc(title)}</h1>
  <div class="meta">${meta.map((m) => `<span>${m}</span>`).join('')}</div>
  <p style="margin:0">${intro}</p>
  <h3>Output</h3>${after}
  ${before ? `<h3>Input</h3>${before}` : ''}
  ${notes ? `<div class="note">${notes}</div>` : ''}
  <h3>Script (excerpt) · <a href="${slug}">full source</a></h3>
  <pre><code>${esc(code)}</code></pre>
</article>
<footer>Portfolio sample · fictional data · Node.js + ExcelJS</footer>
</body>
</html>`;
}

const excerpt = (file, from, lines) => fs.readFileSync(file, 'utf8').split('\n').slice(from, from + lines).join('\n');

async function main() {
  const A = __dirname;

  // 1. Sales report
  const s = path.join(A, 'sales-report');
  const raw = fs.readFileSync(path.join(s, 'data', '2026-07.csv'), 'utf8').split('\n').slice(0, 6);
  fs.writeFileSync(path.join(s, 'index.html'), page({
    slug: 'merge-sales.js', title: 'Monthly sales CSVs → Excel report', kicker: 'Excel automation · merge & summarize',
    meta: ['<b>Input:</b> 3 monthly CSV exports', '<b>Output:</b> sales-report.xlsx (3 sheets)', '<b>Time saved:</b> ~1 hour/month of copy-paste'],
    intro: 'Every month the shop exported orders to a new CSV and someone merged them by hand. This script reads every CSV in a folder, calculates revenue, and writes a formatted workbook with a region × month summary (live SUM formulas), a top-products ranking and a filterable sheet with all orders.',
    before: `<div class="flow"><span>2026-07.csv</span><span>2026-08.csv</span><span>2026-09.csv</span> → <span>merge-sales.js</span> → <span>sales-report.xlsx</span></div><pre><code>${esc(raw.join('\n'))}\n…</code></pre>`,
    after: (await sheetTable(path.join(s, 'sales-report.xlsx'), 'Summary')) + (await sheetTable(path.join(s, 'sales-report.xlsx'), 'Top products')),
    code: excerpt(path.join(s, 'merge-sales.js'), 26, 22),
  }));

  // 2. Clean customers
  const c = path.join(A, 'clean-customers');
  const rawC = fs.readFileSync(path.join(c, 'customers-raw.csv'), 'utf8');
  fs.writeFileSync(path.join(c, 'index.html'), page({
    slug: 'clean-customers.js', title: 'Messy customer list → clean Excel', kicker: 'Data cleaning · deduplication',
    meta: ['<b>Input:</b> customers-raw.csv', '<b>Output:</b> customers-clean.xlsx + Issues sheet', '<b>Fixes:</b> names, emails, phones, cities, dates, duplicates'],
    intro: 'A CRM export with mixed casing, three date formats, phones written five different ways and repeated customers. The script normalizes every field, removes duplicates by email, highlights rows that still need a human look, and lists every problem it found.',
    before: `<pre><code>${esc(rawC)}</code></pre>`,
    after: (await sheetTable(path.join(c, 'customers-clean.xlsx'), 'Customers')) + (await sheetTable(path.join(c, 'customers-clean.xlsx'), 'Issues')),
    code: excerpt(path.join(c, 'clean-customers.js'), 12, 18),
  }));

  // 3. Invoices
  const v = path.join(A, 'invoices');
  const pdfs = fs.readdirSync(path.join(v, 'pdf'));
  fs.writeFileSync(path.join(v, 'index.html'), page({
    slug: 'make-invoices.js', title: 'Excel billing sheet → PDF invoices', kicker: 'Document automation · PDF',
    meta: ['<b>Input:</b> billing.xlsx', `<b>Output:</b> ${pdfs.length} PDF invoices`, '<b>Grouping:</b> one invoice per client'],
    intro: 'Billable lines are kept in a simple spreadsheet. The script groups them by client, numbers the invoices, calculates line amounts and totals, and creates a clean branded PDF for each client, ready to email.',
    before: await sheetTable(path.join(v, 'billing.xlsx')),
    after: `<table><tr><th>Generated file</th><th>Preview</th></tr>${pdfs.map((f) => `<tr><td>${esc(f)}</td><td><a href="pdf/${encodeURIComponent(f)}">open PDF</a></td></tr>`).join('')}</table>`,
    code: excerpt(path.join(v, 'make-invoices.js'), 34, 24),
  }));

  // 4. Scraping
  const b = path.join(A, 'scrape-books');
  fs.writeFileSync(path.join(b, 'index.html'), page({
    slug: 'scrape-books.js', title: 'Website catalogue → Excel (web scraping)', kicker: 'Web scraping · data collection',
    meta: ['<b>Source:</b> books.toscrape.com (public practice site)', '<b>Output:</b> books.xlsx with Stats sheet', '<b>Polite:</b> 0.5 s delay between pages'],
    intro: 'Collects title, price, star rating, stock and link for every product across several catalogue pages, then saves a filterable Excel sheet plus quick stats. The same approach works for price monitoring, lead lists or product research on sites that allow it.',
    before: '',
    after: (await sheetTable(path.join(b, 'books.xlsx'), 'Books', 10, [0, 1, 2, 3])) + (await sheetTable(path.join(b, 'books.xlsx'), 'Stats')),
    code: excerpt(path.join(b, 'scrape-books.js'), 11, 22),
  }));
  console.log('pages built');
}
main();
