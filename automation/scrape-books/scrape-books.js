// Scrape a product catalogue into Excel (books.toscrape.com is a public site made for scraping practice).
// Usage: node scrape-books.js <pages> <output.xlsx>
const ExcelJS = require('exceljs');
const cheerio = require('cheerio');

const [pages = '3', out = 'books.xlsx'] = process.argv.slice(2);
const BASE = 'https://books.toscrape.com/catalogue/';
const RATING = { One: 1, Two: 2, Three: 3, Four: 4, Five: 5 };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const books = [];
  for (let p = 1; p <= Number(pages); p++) {
    const html = await (await fetch(`${BASE}page-${p}.html`)).text();
    const $ = cheerio.load(html);
    $('article.product_pod').each((_, el) => {
      const a = $(el).find('h3 a');
      books.push({
        title: a.attr('title'),
        price_gbp: Number($(el).find('.price_color').text().replace(/[^\d.]/g, '')),
        rating: RATING[$(el).find('.star-rating').attr('class').split(' ')[1]],
        in_stock: $(el).find('.instock').text().includes('In stock') ? 'Yes' : 'No',
        url: BASE + a.attr('href'),
      });
    });
    await sleep(500); // be polite to the server
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Books');
  ws.columns = [{ header: 'Title', key: 'title', width: 50 }, { header: 'Price (£)', key: 'price_gbp', width: 11 }, { header: 'Rating', key: 'rating', width: 8 }, { header: 'In stock', key: 'in_stock', width: 10 }, { header: 'URL', key: 'url', width: 60 }];
  books.forEach((b) => ws.addRow(b));
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } };
  ws.autoFilter = { from: 'A1', to: 'E1' };
  const st = wb.addWorksheet('Stats');
  st.addRows([['Books scraped', books.length], ['Average price (£)', +(books.reduce((s, b) => s + b.price_gbp, 0) / books.length).toFixed(2)], ['5-star books', books.filter((b) => b.rating === 5).length], ['Cheapest', books.reduce((a, b) => (b.price_gbp < a.price_gbp ? b : a)).title]]);
  st.getColumn(1).width = 22; st.getColumn(2).width = 40;
  await wb.xlsx.writeFile(out);
  console.log(`${books.length} books from ${pages} pages -> ${out}`);
}
main();
