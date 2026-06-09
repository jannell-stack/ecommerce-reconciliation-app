const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const TEMPLATE_STYLES_PATH = path.join(
  __dirname,
  '..',
  '..',
  'templates',
  'ESK-Payment-template.xlsx',
);

const DATE_SERIAL_MIN = 30000;
const DATE_SERIAL_MAX = 60000;
const TEMPLATE_DATE_STYLE_INDEX = '14';

let cachedTemplateStylesXml = null;

function isExcelDateSerial(value) {
  const n = Number(value);
  return Number.isFinite(n) && n >= DATE_SERIAL_MIN && n <= DATE_SERIAL_MAX;
}

function loadTemplateStylesXml() {
  if (cachedTemplateStylesXml) return cachedTemplateStylesXml;

  const buffer = fs.readFileSync(TEMPLATE_STYLES_PATH);
  return JSZip.loadAsync(buffer).then(async (zip) => {
    const stylesXml = await zip.file('xl/styles.xml')?.async('string');
    if (!stylesXml) {
      throw new Error('Template styles.xml is missing from ESK-Payment-template.xlsx');
    }
    cachedTemplateStylesXml = stylesXml;
    return stylesXml;
  });
}

function patchSheetColumnBDateStyles(sheetXml, styleIndex) {
  return sheetXml.replace(
    /<c r="(B\d+)"([^>]*)>((?:(?!<\/c>).)*)<\/c>/g,
    (match, address, attrs, inner) => {
      const valueMatch = inner.match(/<v>(\d+)<\/v>/);
      if (!valueMatch || !isExcelDateSerial(Number(valueMatch[1]))) {
        return match;
      }
      const cleanedAttrs = attrs.replace(/\s+s="[^"]*"/, '');
      return `<c r="${address}" s="${styleIndex}"${cleanedAttrs}>${inner}</c>`;
    },
  );
}

async function applyColumnBDateStyles(buffer) {
  const templateStylesXml = await loadTemplateStylesXml();
  const zip = await JSZip.loadAsync(buffer);
  const sheetFile = zip.file('xl/worksheets/sheet1.xml');

  if (!sheetFile) {
    return buffer;
  }

  const sheetXml = await sheetFile.async('string');
  const patchedSheetXml = patchSheetColumnBDateStyles(
    sheetXml,
    TEMPLATE_DATE_STYLE_INDEX,
  );

  zip.file('xl/styles.xml', templateStylesXml);
  zip.file('xl/worksheets/sheet1.xml', patchedSheetXml);

  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

module.exports = {
  applyColumnBDateStyles,
  isExcelDateSerial,
};
