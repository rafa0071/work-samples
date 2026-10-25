// Clean a messy customer CSV: fix names, emails, phones, cities and dates,
// remove duplicates and list every problem found in an "Issues" sheet.
// Usage: node clean-customers.js <customers-raw.csv> <customers-clean.xlsx>
const fs = require('fs');
const ExcelJS = require('exceljs');

const [input = 'customers-raw.csv', out = 'customers-clean.xlsx'] = process.argv.slice(2);
const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const CITY_FIX = { bogota: 'Bogotá', 'sao paulo': 'São Paulo', 'new york': 'New York', paris: 'Paris', riyadh: 'Riyadh', shanghai: 'Shanghai' };

const titleCase = (s) => s.trim().replace(/\s+/g, ' ').toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
const normCity = (s) => CITY_FIX[s.trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')] || titleCase(s);

function normPhone(s) {
  const digits = s.replace(/\D/g, '').replace(/^1(?=\d{10}$)/, '');
  return digits.length === 10 ? `+1 ${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}` : null;
}
function normDate(s) {
  s = s.trim();
  let m;
  if ((m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/))) return [m[1], m[2], m[3]];
  if ((m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/))) return [m[3], m[1], m[2]]; // US style MM-DD-YYYY
  if ((m = s.match(/^(\d{1,2}) (\w{3}) (\d{4})$/))) return [m[3], MONTHS[m[2].toLowerCase()], m[1]];
  return null;
}

async function main() {
  const [header, ...lines] = fs.readFileSync(input, 'utf8').trim().split(/\r?\n/);
  const cols = header.split(',');
  const clean = [], issues = [], seen = new Map();

  lines.forEach((line, i) => {
    const rowNo = i + 2;
    const r = Object.fromEntries(cols.map((c, j) => [c, (line.split(',')[j] || '')]));
    if (!r.name.trim()) { issues.push([rowNo, 'name', 'Empty name, row skipped']); return; }
    const email = r.email.trim().toLowerCase();
    const emailOk = /^[^@\s]+@[^@\s]+\.[a-z]{2,}$/.test(email);
    if (!emailOk) issues.push([rowNo, 'email', `Invalid email "${r.email.trim()}"`]);
    const phone = normPhone(r.phone);
    if (!phone) issues.push([rowNo, 'phone', `Could not read phone "${r.phone}"`]);
    const d = normDate(r.signup_date);
    if (!d) issues.push([rowNo, 'signup_date', `Unknown date "${r.signup_date}"`]);
    const rec = { name: titleCase(r.name), email, phone: phone || '', city: normCity(r.city), signup_date: d ? `${d[0]}-${String(d[1]).padStart(2, '0')}-${String(d[2]).padStart(2, '0')}` : '', status: emailOk ? 'OK' : 'Check email' };
    if (seen.has(email)) { issues.push([rowNo, 'duplicate', `Duplicate of row ${seen.get(email)} (${rec.name})`]); return; }
    seen.set(email, rowNo);
    clean.push(rec);
  });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Customers');
  ws.columns = ['name', 'email', 'phone', 'city', 'signup_date', 'status'].map((k) => ({ header: k, key: k, width: k === 'email' ? 30 : 18 }));
  clean.forEach((r) => {
    const row = ws.addRow(r);
    if (r.status !== 'OK') row.getCell('status').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE2DD' } };
  });
  const is = wb.addWorksheet('Issues');
  is.columns = [{ header: 'source row', width: 12 }, { header: 'field', width: 14 }, { header: 'problem', width: 50 }];
  issues.forEach((x) => is.addRow(x));
  for (const s of wb.worksheets) { s.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }; s.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } }; }
  await wb.xlsx.writeFile(out);
  console.log(`${lines.length} raw rows -> ${clean.length} clean customers, ${issues.length} issues -> ${out}`);
}
main();
