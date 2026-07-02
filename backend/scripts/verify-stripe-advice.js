const fs = require('fs');
const XLSX = require('xlsx');
const { createStripeAdviceWorkbook } = require('../src/services/stripeAdvice');

const stripePath = 'C:/Users/jannell.lim/Desktop/Mar_03.Stripe Mar26.xlsx';
const ncPath = 'C:/Users/jannell.lim/Desktop/Mar_NC Payment Listing-Detail - Mar26.xlsx';
const referencePath = stripePath;

const generated = createStripeAdviceWorkbook(
  fs.readFileSync(stripePath),
  fs.readFileSync(ncPath),
  {
    stripeFileName: 'Mar_03.Stripe Mar26.xlsx',
    ncFileName: 'Mar_NC Payment Listing-Detail - Mar26.xlsx',
  },
);

const refWb = XLSX.readFile(referencePath);
const refRows = XLSX.utils.sheet_to_json(refWb.Sheets['Stripe Advice Mar26'], {
  header: 1,
  defval: '',
});
const genWb = XLSX.read(generated.buffer, { type: 'buffer' });
const genRows = XLSX.utils.sheet_to_json(genWb.Sheets[generated.sheetName], {
  header: 1,
  defval: '',
});

console.log('generated rows', generated.generated, 'linked', generated.linked);
console.log('reference rows', refRows.length - 1);
console.log('headers match', JSON.stringify(refRows[0]) === JSON.stringify(genRows[0]));
console.log('gen headers', genRows[0]);

function rowKey(row) {
  return [row[1], row[3], row[5], row[2]].join('|');
}

const refMap = new Map(refRows.slice(1).map((row) => [rowKey(row), row]));
let fullMatch = 0;
let partial = 0;

for (const row of genRows.slice(1)) {
  const ref = refMap.get(rowKey(row));
  if (!ref) continue;
  partial += 1;
  if (JSON.stringify(row.slice(0, 9)) === JSON.stringify(ref.slice(0, 9))) {
    fullMatch += 1;
  }
}

console.log('matched keys', partial, 'full row match', fullMatch);

const sample = genRows[2];
const refSample = refRows.find(
  (row) => row[1] === 'charge' && String(row[3]) === String(sample[3]),
);
console.log('sample generated', sample);
console.log('sample reference', refSample);
