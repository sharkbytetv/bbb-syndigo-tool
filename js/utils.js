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

async function loadTemplate(path) {
  const resp = await fetch(path);
  if (!resp.ok) throw new Error(`Failed to load template: ${path} (${resp.status})`);
  const buf = await resp.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
  for (const name of wb.SheetNames) {
    delete wb.Sheets[name]['!dataValidation'];
    delete wb.Sheets[name]['!dvs'];
  }
  return wb;
}

// Appends dataRows after the existing header row in the named sheet.
// Preserves the header, hidden sheets, formatting, and dropdown lists from the template.
function fillSheet(wb, sheetName, dataRows) {
  if (!wb.Sheets[sheetName]) return;
  const ws = wb.Sheets[sheetName];
  if (dataRows.length === 0) return;
  XLSX.utils.sheet_add_aoa(ws, dataRows, { origin: { r: 1, c: 0 } });
  // Extend the sheet's declared range to cover the new rows
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const lastRow = dataRows.length; // 0-indexed row 1 = Excel row 2, so last = dataRows.length
  const lastCol = Math.max(range.e.c, Math.max(...dataRows.map(r => r.length - 1)));
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: lastCol } });
}

// Updates TENANT and DOMAIN in METADATA by searching column A for the label.
function updateMetadata(wb, tenant, domain) {
  const ws = wb.Sheets['METADATA'];
  if (!ws) return;
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  aoa.forEach((row, ri) => {
    const label = row[0] ? String(row[0]).trim() : '';
    if (label === 'TENANT') {
      const addr = XLSX.utils.encode_cell({ r: ri, c: 1 });
      ws[addr] = { v: tenant, t: 's' };
    }
    if (label === 'DOMAIN') {
      const addr = XLSX.utils.encode_cell({ r: ri, c: 1 });
      ws[addr] = { v: domain, t: 's' };
    }
  });
}
