const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const inputFile = 'ПРАЙС т,вода 24,10,26.xls';
const outputFile = 'products.json';

const parsePrice = (value) => {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  const str = String(value).replace(/\s+/g, '').replace(',', '.');
  const num = parseFloat(str);
  return isNaN(num) ? null : num;
}

const parseBarcode = (value) => {
  if (value === null || value === undefined) return null;
  const str = String(value).trim();
  if (!str) return null;
  // barcode usually digits, may have length >= 8
  if (/^\d{8,}$/.test(str)) return str;
  return null;
}

const main = () => {
  try {
    const workbook = XLSX.readFile(inputFile);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    const products = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!Array.isArray(row)) continue;
      
      const barcode = parseBarcode(row[7]);
      const name = String(row[1] || '').trim();
      const price = parsePrice(row[3]);
      
      // Skip if missing required fields
      if (!barcode || !name || price === null) {
        continue;
      }
      
      // Skip if name looks like section header
      if (name.toLowerCase().includes('парфюмерия') && row[7] === '') {
        continue;
      }
      
      products.push({
        barcode: barcode,
        name: name,
        price: price
      });
    }
    
    fs.writeFileSync(outputFile, JSON.stringify(products, null, 2), 'utf8');
    console.log(`Parsed ${products.length} products`);
    console.log(`Saved to ${outputFile}`);
  } catch (error) {
    console.error('Error:', error.message);
    process.exit(1);
  }
}

main();
