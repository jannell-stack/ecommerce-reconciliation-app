const fs = require('fs');
const path = require('path');
const os = require('os');
const JSZip = require('jszip');
const XLSX = require('xlsx');

const SK = 'c:/Users/jannell.lim/Desktop/SK Payment Verification Combined Y2026.xlsx';
const NC = 'c:/Users/jannell.lim/Desktop/NC Payment Listing-Detail - May26.xlsx';
const TPL = 'templates/ESK-Payment-template.xlsx';

function extractColumnBDateCells(sheetXml) {
  const cells = [];
  for (const match of sheetXml.matchAll(/<c r="(B\d+)"([^>]*)>((?:(?!<\/c>).)*)<\/c>/g)) {
    const valueMatch = match[3].match(/<v>(\d+)<\/v>/);
    if (!valueMatch) continue;
    const value = Number(valueMatch[1]);
    if (value < 30000 || value > 60000) continue;
    cells.push({
      address: match[1],
      style: match[2].match(/\ss="(\d+)"/)?.[1] || null,
      value,
    });
  }
  return cells;
}

async function auditBuffer(buffer, label) {
  const zip = await JSZip.loadAsync(buffer);
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('string') || '';
  const cells = extractColumnBDateCells(sheetXml);
  const wrong = cells.filter((cell) => cell.style !== '14' && cell.style !== '13');
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true, cellNF: true });
  const ws = wb.Sheets[wb.SheetNames[0]];

  console.log(`\n${label}`);
  console.log('  zip:', buffer.subarray(0, 2).toString() === 'PK' ? 'OK' : 'BAD');
  console.log('  date cells:', cells.length, 'wrong style:', wrong.length);
  for (const row of [3, 64, 221]) {
    const cell = cells.find((c) => c.address === `B${row}`);
    console.log(`  B${row}: style=${cell?.style} display=${ws[`B${row}`]?.w}`);
  }

  return wrong.length === 0 && buffer.subarray(0, 2).toString() === 'PK';
}

async function main() {
  const sk = fs.readFileSync(SK);
  const nc = fs.readFileSync(NC);

  for (const mode of ['dev', 'production']) {
    delete process.env.ESK_VALUES_ONLY;
    delete process.env.NODE_ENV;
    if (mode === 'production') process.env.ESK_VALUES_ONLY = 'true';
    delete require.cache[require.resolve('../src/services/eskPayment')];
    const { createEskPaymentWorkbook } = require('../src/services/eskPayment');
    const result = await createEskPaymentWorkbook(sk, nc);
    const ok = await auditBuffer(result.buffer, mode.toUpperCase());
    if (!ok) {
      console.error(`\n${mode} FAILED`);
      process.exit(1);
    }
    if (mode === 'dev') {
      const out = path.join(os.tmpdir(), 'ESK-Payment-May26-VERIFIED.xlsx');
      fs.writeFileSync(out, result.buffer);
      console.log('\nVerified dev output:', out);
    }
  }

  const tplStyles = await JSZip.loadAsync(fs.readFileSync(TPL))
    .then((z) => z.file('xl/styles.xml').async('string'));
  const devBuf = fs.readFileSync(path.join(os.tmpdir(), 'ESK-Payment-May26-VERIFIED.xlsx'));
  const outStyles = await JSZip.loadAsync(devBuf)
    .then((z) => z.file('xl/styles.xml').async('string'));
  console.log('\nstyles.xml matches template:', tplStyles === outStyles);
  console.log('\nALL CHECKS PASSED');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
