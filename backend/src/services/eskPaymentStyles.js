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
const GPAY_STYLES_PATH = path.join(__dirname, '..', '..', 'templates', 'GPay-styles.xml');

const DATE_SERIAL_MIN = 30000;
const DATE_SERIAL_MAX = 60000;
const TEMPLATE_DATE_STYLE_INDEX = '14';
const TEMPLATE_GENERAL_STYLE_INDEX = '0';
const GPAY_DATE_STYLE_INDEX = '2';
const GPAY_NUMBER_STYLE_INDEX = '4';
const GPAY_DIFF_STYLE_INDEX = '9';

let cachedTemplateStylesXml = null;
let cachedGpayStylesXml = null;

function isExcelDateSerial(value) {
  const n = Number(value);
  return Number.isFinite(n) && n >= DATE_SERIAL_MIN && n <= DATE_SERIAL_MAX;
}

function loadTemplateStylesXml() {
  if (cachedTemplateStylesXml) return Promise.resolve(cachedTemplateStylesXml);

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

function loadGpayStylesXml() {
  if (cachedGpayStylesXml) return Promise.resolve(cachedGpayStylesXml);

  if (!fs.existsSync(GPAY_STYLES_PATH)) {
    throw new Error('GPay-styles.xml is missing from backend/templates');
  }

  cachedGpayStylesXml = fs.readFileSync(GPAY_STYLES_PATH, 'utf8');
  return Promise.resolve(cachedGpayStylesXml);
}

function appendBorderlessNumberStyle(stylesXml) {
  const cellXfsMatch = stylesXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  if (!cellXfsMatch) {
    return { stylesXml, styleIndex: TEMPLATE_GENERAL_STYLE_INDEX };
  }

  const count = Number(cellXfsMatch[1]);
  const newXf =
    '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf>';
  return {
    stylesXml: stylesXml.replace(
      cellXfsMatch[0],
      `<cellXfs count="${count + 1}">${cellXfsMatch[2]}${newXf}</cellXfs>`,
    ),
    styleIndex: String(count),
  };
}

function appendBorderlessDateStyle(stylesXml) {
  const cellXfsMatch = stylesXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  if (!cellXfsMatch) {
    return { stylesXml, styleIndex: TEMPLATE_DATE_STYLE_INDEX };
  }

  const count = Number(cellXfsMatch[1]);
  const newXf =
    '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center"/></xf>';
  return {
    stylesXml: stylesXml.replace(
      cellXfsMatch[0],
      `<cellXfs count="${count + 1}">${cellXfsMatch[2]}${newXf}</cellXfs>`,
    ),
    styleIndex: String(count),
  };
}

function patchSheetColumnCells(sheetXml, columnLetter, styleIndex, predicate) {
  return sheetXml.replace(
    new RegExp(`<c r="(${columnLetter}\\d+)"([^>]*)>((?:(?!<\\/c>).)*)<\\/c>`, 'g'),
    (match, address, attrs, inner) => {
      if (predicate && !predicate(address, inner)) {
        return match;
      }
      const cleanedAttrs = attrs.replace(/\s+s="[^"]*"/, '');
      return `<c r="${address}" s="${styleIndex}"${cleanedAttrs}>${inner}</c>`;
    },
  );
}

function patchSheetColumnDateStyles(sheetXml, columnLetter, styleIndex) {
  return patchSheetColumnCells(sheetXml, columnLetter, styleIndex, (_address, inner) => {
    if (/<f>/.test(inner)) return true;
    const valueMatch = inner.match(/<v>(\d+(?:\.\d+)?)<\/v>/);
    return Boolean(valueMatch && isExcelDateSerial(Number(valueMatch[1])));
  });
}

function patchSheetHeaderRowStyles(sheetXml, columns, rowNumber, styleIndex) {
  let patched = sheetXml;
  for (const columnLetter of columns) {
    patched = patchSheetColumnCells(
      patched,
      columnLetter,
      styleIndex,
      (address) => address === `${columnLetter}${rowNumber}`,
    );
  }
  return patched;
}

async function applyWorkbookStyles(
  buffer,
  {
    sheetPatches = [],
    replaceStyles = true,
    stylesLoader = loadTemplateStylesXml,
    appendCustomStyles = true,
    dateStyleIndex: presetDateStyleIndex,
    numberStyleIndex: presetNumberStyleIndex,
    generalStyleIndex = TEMPLATE_GENERAL_STYLE_INDEX,
  } = {},
) {
  let stylesXml = replaceStyles ? await stylesLoader() : null;
  let numberStyleIndex = presetNumberStyleIndex ?? TEMPLATE_GENERAL_STYLE_INDEX;
  let dateStyleIndex = presetDateStyleIndex ?? TEMPLATE_DATE_STYLE_INDEX;

  if (replaceStyles && appendCustomStyles) {
    const borderlessNumberStyle = appendBorderlessNumberStyle(stylesXml);
    stylesXml = borderlessNumberStyle.stylesXml;
    numberStyleIndex = borderlessNumberStyle.styleIndex;

    const borderlessDateStyle = appendBorderlessDateStyle(stylesXml);
    stylesXml = borderlessDateStyle.stylesXml;
    dateStyleIndex = borderlessDateStyle.styleIndex;
  }

  const zip = await JSZip.loadAsync(buffer);

  for (const patch of sheetPatches) {
    const sheetFile = zip.file(patch.sheetPath);
    if (!sheetFile) continue;

    let sheetXml = await sheetFile.async('string');
    for (const columnLetter of patch.numberColumns || []) {
      sheetXml = patchSheetColumnCells(
        sheetXml,
        columnLetter,
        numberStyleIndex,
        (_address, inner) => /<v>/.test(inner) || /<f>/.test(inner),
      );
    }
    for (const [columnLetter, styleIndex] of Object.entries(patch.columnStyleOverrides || {})) {
      sheetXml = patchSheetColumnCells(
        sheetXml,
        columnLetter,
        styleIndex,
        (_address, inner) => /<v>/.test(inner) || /<f>/.test(inner),
      );
    }
    for (const columnLetter of patch.dateColumns || []) {
      sheetXml = patchSheetColumnDateStyles(sheetXml, columnLetter, dateStyleIndex);
    }
    for (const columnLetter of patch.generalColumns || []) {
      sheetXml = patchSheetColumnCells(
        sheetXml,
        columnLetter,
        generalStyleIndex,
        (_address, inner) => /<v>/.test(inner),
      );
    }
    if (patch.headerRow && patch.headerColumns?.length) {
      sheetXml = patchSheetHeaderRowStyles(
        sheetXml,
        patch.headerColumns,
        patch.headerRow,
        generalStyleIndex,
      );
    }
    zip.file(patch.sheetPath, sheetXml);
  }

  if (replaceStyles) {
    zip.file('xl/styles.xml', stylesXml);
  }

  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

async function applyColumnDateStyles(
  buffer,
  { sheetPath = 'xl/worksheets/sheet1.xml', columnLetter = 'C' } = {},
) {
  return applyWorkbookStyles(buffer, {
    sheetPatches: [{ sheetPath, dateColumns: [columnLetter] }],
  });
}

async function applyStripeAdviceWorkbookStyles(buffer) {
  return applyWorkbookStyles(buffer, {
    sheetPatches: [
      {
        sheetPath: 'xl/worksheets/sheet1.xml',
        dateColumns: ['A', 'F'],
        numberColumns: ['C', 'D', 'E', 'G', 'H'],
        generalColumns: ['B'],
        headerRow: 3,
        headerColumns: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
      },
      {
        sheetPath: 'xl/worksheets/sheet2.xml',
        dateColumns: ['C'],
      },
    ],
  });
}

async function applyGpayWorkbookStyles(buffer) {
  return applyWorkbookStyles(buffer, {
    stylesLoader: loadGpayStylesXml,
    appendCustomStyles: false,
    dateStyleIndex: GPAY_DATE_STYLE_INDEX,
    numberStyleIndex: GPAY_NUMBER_STYLE_INDEX,
    sheetPatches: [
      {
        sheetPath: 'xl/worksheets/sheet1.xml',
        dateColumns: ['M'],
        numberColumns: ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'N'],
        columnStyleOverrides: { O: GPAY_DIFF_STYLE_INDEX },
        generalColumns: ['A', 'B', 'C', 'D', 'E'],
        headerRow: 1,
        headerColumns: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'],
      },
      {
        sheetPath: 'xl/worksheets/sheet2.xml',
        dateColumns: ['C'],
      },
    ],
  });
}

async function applyShopbackWorkbookStyles(buffer) {
  return applyWorkbookStyles(buffer, {
    stylesLoader: loadGpayStylesXml,
    appendCustomStyles: false,
    dateStyleIndex: GPAY_DATE_STYLE_INDEX,
    numberStyleIndex: GPAY_NUMBER_STYLE_INDEX,
    sheetPatches: [
      {
        sheetPath: 'xl/worksheets/sheet1.xml',
        dateColumns: ['C', 'J'],
        numberColumns: ['E', 'F', 'G', 'H', 'I', 'K', 'L'],
        generalColumns: ['A', 'B', 'D', 'M'],
        headerRow: 1,
        headerColumns: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'],
      },
      {
        sheetPath: 'xl/worksheets/sheet2.xml',
        dateColumns: ['C'],
      },
    ],
  });
}

async function applyColumnBDateStyles(buffer) {
  return applyColumnDateStyles(buffer, {
    sheetPath: 'xl/worksheets/sheet1.xml',
    columnLetter: 'C',
  });
}

module.exports = {
  applyColumnBDateStyles,
  applyColumnDateStyles,
  applyGpayWorkbookStyles,
  applyShopbackWorkbookStyles,
  applyStripeAdviceWorkbookStyles,
  isExcelDateSerial,
};
