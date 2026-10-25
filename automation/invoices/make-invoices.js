// Turn an Excel sheet of billable lines into one PDF invoice per client.
// Usage: node make-invoices.js <billing.xlsx> <output-folder>
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');

const [input = 'billing.xlsx', outDir = 'pdf'] = process.argv.slice(2);
const money = (n) => '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

async function createSampleWorkbook(file) {
  // Fictional billing data so the sample is reproducible
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Billing');
  ws.columns = ['client', 'email', 'description', 'qty', 'rate'].map((k) => ({ header: k, key: k, width: 26 }));
  [
    ['Harbor Coffee Roasters', 'ap@harborcoffee.example', 'Website maintenance (October)', 1, 180],
    ['Harbor Coffee Roasters', 'ap@harborcoffee.example', 'Product photo retouching', 12, 8.5],
    ['Lumen Studio', 'hello@lumenstudio.example', 'Booking page updates', 3, 45],
    ['Lumen Studio', 'hello@lumenstudio.example', 'Monthly analytics report', 1, 90],
    ['Northwind Goods', 'billing@northwind.example', 'Data cleanup script', 1, 240],
  ].forEach((r) => ws.addRow(r));
  await wb.xlsx.writeFile(file);
}

async function main() {
  if (!fs.existsSync(input)) await createSampleWorkbook(input);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(input);
  const ws = wb.worksheets[0];
  const keys = ws.getRow(1).values.slice(1);
  const lines = [];
  ws.eachRow((row, i) => { if (i > 1) lines.push(Object.fromEntries(keys.map((k, j) => [k, row.values[j + 1]]))); });

  const clients = [...new Set(lines.map((l) => l.client))];
  fs.mkdirSync(outDir, { recursive: true });
  let n = 140;
  for (const client of clients) {
    const items = lines.filter((l) => l.client === client);
    const number = `INV-2026-${++n}`;
    const doc = new PDFDocument({ size: 'A4', margin: 50 });
    const file = path.join(outDir, `${number} ${client}.pdf`);
    doc.pipe(fs.createWriteStream(file));
    doc.fillColor('#1F7A4D').fontSize(26).text('INVOICE', { align: 'right' });
    doc.fillColor('#333').fontSize(10).text(number, { align: 'right' }).text('Issued: Oct 25, 2026 · Due in 14 days', { align: 'right' });
    doc.moveDown(2).fontSize(11).fillColor('#888').text('BILL TO').fillColor('#111').fontSize(13).text(client).fontSize(10).fillColor('#555').text(items[0].email);
    doc.moveDown(1.5);
    const top = doc.y;
    doc.fontSize(9).fillColor('#888').text('DESCRIPTION', 50, top).text('QTY', 330, top).text('RATE', 390, top).text('AMOUNT', 470, top);
    doc.moveTo(50, top + 14).lineTo(545, top + 14).strokeColor('#ddd').stroke();
    let y = top + 22, total = 0;
    for (const it of items) {
      const amount = it.qty * it.rate;
      total += amount;
      doc.fontSize(10).fillColor('#111').text(it.description, 50, y, { width: 270 }).text(String(it.qty), 330, y).text(money(it.rate), 390, y).text(money(amount), 470, y);
      y += 22;
    }
    doc.moveTo(330, y + 4).lineTo(545, y + 4).strokeColor('#1F7A4D').stroke();
    doc.fontSize(13).fillColor('#1F7A4D').text('Total ' + money(total), 330, y + 12, { width: 215, align: 'right' });
    doc.fontSize(9).fillColor('#888').text('Thank you for your business.', 50, 760);
    doc.end();
    console.log(`${file}  ${items.length} lines  ${money(total)}`);
  }
}
main();
