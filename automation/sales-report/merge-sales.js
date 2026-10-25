// Merge monthly sales CSV exports into one formatted Excel report.
// Usage: node merge-sales.js <folder-with-csvs> <output.xlsx>
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const [dir = 'data', out = 'sales-report.xlsx'] = process.argv.slice(2);

function readCsv(file) {
  const [header, ...lines] = fs.readFileSync(file, 'utf8').trim().split(/\r?\n/);
  const cols = header.split(',');
  return lines.map((line) => {
    const cells = line.split(',');
    const row = Object.fromEntries(cols.map((c, i) => [c.trim(), cells[i]?.trim()]));
    row.units = Number(row.units);
    row.unit_price = Number(row.unit_price);
    row.revenue = +(row.units * row.unit_price).toFixed(2);
    row.month = path.basename(file, '.csv');
    return row;
  });
}

async function main() {
  const files = fs.readdirSync(dir).filter((f) => f.endsWith('.csv')).sort();
  const rows = files.flatMap((f) => readCsv(path.join(dir, f)));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'sales-report script';

  // Summary: revenue per region and month, with totals
  const months = [...new Set(rows.map((r) => r.month))];
  const regions = [...new Set(rows.map((r) => r.region))].sort();
  const sum = wb.addWorksheet('Summary', { views: [{ state: 'frozen', ySplit: 1 }] });
  sum.columns = [{ header: 'Region', key: 'region', width: 16 }, ...months.map((m) => ({ header: m, key: m, width: 14 })), { header: 'Total', key: 'total', width: 16 }];
  for (const region of regions) {
    const rec = { region };
    for (const m of months) rec[m] = rows.filter((r) => r.region === region && r.month === m).reduce((s, r) => s + r.revenue, 0);
    rec.total = months.reduce((s, m) => s + rec[m], 0);
    sum.addRow(rec);
  }
  const last = sum.rowCount;
  const totalRow = sum.addRow({ region: 'Total' });
  [...months, 'total'].forEach((key, i) => {
    const col = sum.getColumn(i + 2).letter;
    totalRow.getCell(key).value = { formula: `SUM(${col}2:${col}${last})` };
  });

  // Top products
  const prod = wb.addWorksheet('Top products');
  prod.columns = [{ header: 'Product', key: 'product', width: 24 }, { header: 'Units', key: 'units', width: 10 }, { header: 'Revenue', key: 'revenue', width: 14 }];
  const byProduct = {};
  for (const r of rows) {
    byProduct[r.product] ??= { product: r.product, units: 0, revenue: 0 };
    byProduct[r.product].units += r.units;
    byProduct[r.product].revenue += r.revenue;
  }
  Object.values(byProduct).sort((a, b) => b.revenue - a.revenue).forEach((p) => prod.addRow(p));

  // All rows, filterable
  const all = wb.addWorksheet('All orders');
  all.columns = ['month', 'order_id', 'region', 'product', 'units', 'unit_price', 'revenue'].map((k) => ({ header: k, key: k, width: k === 'product' ? 24 : 13 }));
  rows.forEach((r) => all.addRow(r));
  all.autoFilter = { from: 'A1', to: 'G1' };

  // Shared formatting
  for (const ws of wb.worksheets) {
    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } };
    ws.eachRow((row, i) => { if (i > 1) row.eachCell((c) => { if (typeof c.value === 'number' || c.value?.formula) c.numFmt = '#,##0.00'; }); });
  }
  totalRow.font = { bold: true };
  await wb.xlsx.writeFile(out);
  console.log(`${rows.length} orders from ${files.length} files -> ${out}`);
}
main();
