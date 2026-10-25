// Generates the fictional input data used by the automation samples.
const fs = require('fs');
const path = require('path');
let seed = 42;
const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
const pick = (a) => a[Math.floor(rnd() * a.length)];

// 1) Monthly sales CSVs
const salesDir = path.join(__dirname, 'sales-report', 'data');
fs.mkdirSync(salesDir, { recursive: true });
const products = [['Desk lamp', 39.9], ['Wool throw', 54], ['Mug set', 24.5], ['Planter', 18], ['Wall clock', 32]];
let id = 1000;
for (const month of ['2026-07', '2026-08', '2026-09']) {
  const lines = ['order_id,region,product,units,unit_price'];
  for (let i = 0; i < 40; i++) {
    const [p, price] = pick(products);
    lines.push([++id, pick(['North', 'South', 'East', 'West']), p, 1 + Math.floor(rnd() * 5), price].join(','));
  }
  fs.writeFileSync(path.join(salesDir, month + '.csv'), lines.join('\n'));
}

// 2) Messy customer list
const cleanDir = path.join(__dirname, 'clean-customers');
fs.mkdirSync(cleanDir, { recursive: true });
const messy = [
  'name,email,phone,city,signup_date',
  '  maria GARCIA ,Maria.Garcia@Example.com ,(555) 010-2233,bogota,2026/03/04',
  'John Smith,john.smith@example.com,555.010.4455,New York,04-03-2026',
  'maria garcia,maria.garcia@example.com,5550102233,Bogotá,2026-03-04',
  'Li Wei,li.wei@example,555-010-7788,Shanghai,2026-03-07',
  'Ana  Souza,ana.souza@example.com,+1 555 010 9900,sao paulo,7 Mar 2026',
  ',empty.row@example.com,,,',
  'JOHN SMITH,JOHN.SMITH@EXAMPLE.COM,555 010 4455,new york,2026-03-04',
  'Omar Haddad,omar.haddad@example.com,555-010-1212,Riyadh,2026-03-09',
  'Chloé Martin,chloe.martin@example.com,555 010 3434,paris,2026-03-10',
];
fs.writeFileSync(path.join(cleanDir, 'customers-raw.csv'), messy.join('\n'));
console.log('sample data written');
