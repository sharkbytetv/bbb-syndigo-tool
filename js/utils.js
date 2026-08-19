const BOOL_MAP = { Y: 'Yes', N: 'No', Yes: 'Yes', No: 'No', TRUE: 'Yes', FALSE: 'No', true: 'Yes', false: 'No' };

function normalizeBoolean(val) {
  if (val === null || val === undefined || val === '') return 'No';
  return BOOL_MAP[String(val).trim()] ?? 'No';
}

function normalizeDataType(val) {
  if (!val) return 'string';
  const v = String(val).trim().toLowerCase();
  if (v === 'list of values') return 'list of values';
  return v;
}

function inferDisplayType(attr) {
  const dt = normalizeDataType(attr.dataType);
  const name = (attr.displayName || '').toLowerCase();
  if (dt === 'path')           return 'path';
  if (dt === 'boolean')        return 'boolean';
  if (dt === 'list of values') return 'referencelist';
  if (dt === 'date')           return 'date';
  if (dt === 'datetime')       return 'datetime';
  if (dt === 'nested')         return 'nestedgrid';
  if (dt === 'deeplynested')   return 'deeplynested';
  if (dt === 'integer' || dt === 'decimal') return 'textbox';
  if (dt === 'string') {
    if (/description|features|benefits|policy|procedure|details/.test(name)) return 'richtexteditor';
    if (attr.maxLength && Number(attr.maxLength) > 500) return 'textarea';
    return 'textbox';
  }
  return 'textbox';
}

function generateUUID() {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });
}

function refTableForAttr(fieldId, refTables) {
  const suffix = fieldId.replace(/^thg/, '').toLowerCase();
  return refTables.find(r => r.name.replace(/^ref/, '').toLowerCase() === suffix) ?? null;
}

function buildAOASheet(wb, sheetName, rows) {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── JSZip helpers ────────────────────────────────────────────────────────────

function _escXml(v) {
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// 0-based column index → Excel column letter (A, B, …, Z, AA, …)
function _colLetter(n) {
  let s = '';
  n++;
  while (n > 0) {
    s = String.fromCharCode(64 + ((n - 1) % 26) + 1) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Excel column letter → 0-based column index (inverse of _colLetter)
function _colIndex(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

// Rewrite <dimension> to span exactly the cells present in the new sheetData.
// Syndigo templates ship a fixed dimension (e.g. "B3:H703"); leaving it in place
// caps how many rows can be written before Excel reports "We found a problem"
// and repairs the file. Recomputing it removes that ceiling.
function _updateDimension(xml, sheetDataXml) {
  let minC = Infinity, maxC = -1, minR = Infinity, maxR = -1;
  for (const m of sheetDataXml.matchAll(/<c r="([A-Z]+)(\d+)"/g)) {
    const c = _colIndex(m[1]);
    const r = Number(m[2]);
    if (c < minC) minC = c;
    if (c > maxC) maxC = c;
    if (r < minR) minR = r;
    if (r > maxR) maxR = r;
  }
  if (maxR < 0) return xml;  // no cells written; leave the template's dimension alone

  const tag = '<dimension ref="' + _colLetter(minC) + minR + ':' + _colLetter(maxC) + maxR + '"/>';
  if (/<dimension\b[^>]*\/>/.test(xml)) {
    return xml.replace(/<dimension\b[^>]*\/>/, () => tag);
  }
  if (/<dimension\b[^>]*>[\s\S]*?<\/dimension>/.test(xml)) {
    return xml.replace(/<dimension\b[^>]*>[\s\S]*?<\/dimension>/, () => tag);
  }
  return xml;  // no dimension element — Excel computes it on open
}

// Resolve a sheet name to its path inside the ZIP (e.g. "xl/worksheets/sheet1.xml")
async function _findSheetPath(zip, sheetName) {
  const wbXml   = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const esc = sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const sm  = new RegExp('name="' + esc + '"[^>]*r:id="([^"]+)"').exec(wbXml)
           || new RegExp('r:id="([^"]+)"[^>]*name="' + esc + '"').exec(wbXml);
  if (!sm) throw new Error('Sheet not found in workbook: ' + sheetName);
  const rm = new RegExp('Id="' + sm[1] + '"[^>]*Target="([^"]+)"').exec(relsXml);
  if (!rm) throw new Error('Relationship not found for sheet: ' + sheetName);
  const target = rm[1];
  return target.startsWith('xl/') ? target : 'xl/' + target;
}

// Fetch a template file and return a JSZip object — all OOXML content preserved as-is.
async function loadTemplate(path) {
  const resp = await fetch(path);
  if (!resp.ok) throw new Error('Failed to load template: ' + path + ' (' + resp.status + ')');
  const buf = await resp.arrayBuffer();
  return JSZip.loadAsync(buf);
}

// Replace the <sheetData> of `sheetName` with new rows, preserving styled header rows.
// headerCount: how many header rows to keep from the template (default 1).
// dataRows: array of arrays; null/undefined cells are omitted; numbers written as <v>; all else as inlineStr.
async function fillSheet(zip, sheetName, dataRows, headerCount = 1) {
  const sheetPath = await _findSheetPath(zip, sheetName);
  let xml = await zip.file(sheetPath).async('string');

  // Extract the first N header rows from existing sheetData
  const sdMatch    = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml);
  const existingSd = sdMatch ? sdMatch[1] : '';
  let headersXml = '';
  for (let n = 1; n <= headerCount; n++) {
    const m = new RegExp('<row\\b[^>]*\\br="' + n + '"[^>]*>[\\s\\S]*?<\\/row>').exec(existingSd);
    if (m) headersXml += m[0];
  }

  // Build new data rows (starting at headerCount+1) as inline strings / bare numbers
  let newRowsXml = '';
  for (let i = 0; i < dataRows.length; i++) {
    const rowNum = i + headerCount + 1;
    const row = dataRows[i];
    let cellsXml = '';
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (v === null || v === undefined) continue;
      const ref = _colLetter(c) + rowNum;
      if (typeof v === 'number') {
        cellsXml += '<c r="' + ref + '"><v>' + v + '</v></c>';
      } else {
        cellsXml += '<c r="' + ref + '" t="inlineStr"><is><t>' + _escXml(String(v)) + '</t></is></c>';
      }
    }
    if (cellsXml) newRowsXml += '<row r="' + rowNum + '">' + cellsXml + '</row>';
  }

  // Use a function so any $ in headersXml/newRowsXml is treated literally
  const newSd = '<sheetData>' + headersXml + newRowsXml + '</sheetData>';
  xml = xml.replace(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/, () => newSd);
  xml = _updateDimension(xml, newSd);

  zip.file(sheetPath, xml);
}

// Update TENANT and DOMAIN cells in the METADATA sheet via shared-string lookup.
async function updateMetadata(zip, tenant, domain) {
  const sheetPath = await _findSheetPath(zip, 'METADATA');

  const ssFile = zip.file('xl/sharedStrings.xml');
  if (!ssFile) return;
  const ssXml = await ssFile.async('string');

  const ssArr = [];
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let siM;
  while ((siM = siRe.exec(ssXml)) !== null) {
    const texts = [];
    const tRe = /<t\b[^>]*>([^<]*)<\/t>/g;
    let tM;
    while ((tM = tRe.exec(siM[1])) !== null) texts.push(tM[1]);
    ssArr.push(texts.join(''));
  }

  let xml = await zip.file(sheetPath).async('string');

  for (const [label, value] of [['TENANT', tenant], ['DOMAIN', domain]]) {
    const idx = ssArr.indexOf(label);
    if (idx < 0) continue;

    xml = xml.replace(/<row\b([^>]*)>([\s\S]*?)<\/row>/g, (fullRow, attrs, content) => {
      if (!new RegExp('<c r="A\\d+"[^>]*t="s"[^>]*><v>' + idx + '<\\/v><\\/c>').test(content)) {
        return fullRow;
      }
      const rnM = attrs.match(/\br="(\d+)"/);
      if (!rnM) return fullRow;
      const rowNum = rnM[1];

      const existingBM = new RegExp('<c r="B' + rowNum + '"([^>]*)>[\\s\\S]*?<\\/c>').exec(content);
      let sAttr = '';
      if (existingBM) {
        const sM = existingBM[1].match(/s="(\d+)"/);
        if (sM) sAttr = ' s="' + sM[1] + '"';
      }

      const newBCell = '<c r="B' + rowNum + '"' + sAttr + ' t="inlineStr"><is><t>' + _escXml(value) + '</t></is></c>';
      const newContent = existingBM
        ? content.replace(existingBM[0], () => newBCell)
        : content + newBCell;

      return '<row' + attrs + '>' + newContent + '</row>';
    });
  }

  zip.file(sheetPath, xml);
}
