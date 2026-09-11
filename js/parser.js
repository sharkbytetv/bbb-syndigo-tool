function parseDataDictionary(workbook, taxName, taxDisplay) {
  const attrs = parseAttributeMaster(workbook);
  const refTables = parseRefLOVMaster(workbook);
  const taxonomy = parseProductTypeIndex(workbook, taxName, taxDisplay);
  const categoryAttrs = parseCategorySpecificAttrs(workbook, attrs);
  return { attributes: attrs, refTables, taxonomy, categoryAttrs };
}

function parseAttributeMaster(wb) {
  const ws = wb.Sheets['Attribute Master'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const SKIP = new Set(['CORE ATTRIBUTES', 'VARIANT ATTRIBUTES', 'ERP ATTRIBUTES', 'ECOMMERCE ATTRIBUTES', 'CUSTOMER INSIGHT', 'FORECASTING & PLANNING']);
  const attrs = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const displayName = r[0] ? String(r[0]).trim() : null;
    const fieldId = r[1] ? String(r[1]).trim() : null;
    if (!fieldId || !displayName) continue;
    if (SKIP.has(displayName.trim().replace(/^\s+/, ''))) continue;

    const refTableRaw = r[7] ? String(r[7]).trim() : null;
    attrs.push({
      displayName,
      fieldId,
      shopifyMapping: r[2] ? String(r[2]).trim() : null,
      erpMapping: r[3] ? String(r[3]).trim() : null,
      dataType: r[4] ? String(r[4]).trim() : 'string',
      description: r[5] ? String(r[5]).trim() : null,
      group: r[6] ? String(r[6]).trim() : null,
      refTable: (refTableRaw && refTableRaw !== '????' && refTableRaw !== '—') ? refTableRaw : null,
      isProductLevel: normalizeBoolean(r[8]),
      isSkuVariant: normalizeBoolean(r[9]),
      isRequired: normalizeBoolean(r[10]),
      isFilterableOnline: normalizeBoolean(r[11]),
      isMultiValued: normalizeBoolean(r[12]),
      isTranslatable: normalizeBoolean(r[13]),
      isSyndicatable: normalizeBoolean(r[14]),
      isMandatory: normalizeBoolean(r[15]),
      minLength: r[16] ?? null,
      maxLength: r[17] ?? null,
      rangeFrom: r[18] ?? null,
      rangeFromInclusive: normalizeBoolean(r[19]),
      rangeTo: r[20] ?? null,
      rangeToInclusive: normalizeBoolean(r[21]),
      precision: r[22] ?? null,
      minPrecision: r[23] ?? null,
      maxPrecision: r[24] ?? null,
      dependentAttribute: r[25] ? String(r[25]).trim() : null,
    });
  }
  return attrs;
}

function parseRefLOVMaster(wb) {
  const ws = wb.Sheets['Reference LOV Master'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const tables = {};
  let currentTable = null;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const cell0 = String(r[0]).trim();
    if (cell0.includes('—')) {
      const dashParts = cell0.split('—');
      currentTable = dashParts[0].trim();
      const displayFromCell = dashParts.slice(1).join('—').trim();
      const displayFallback = r[1] ? String(r[1]).trim() : currentTable;
      if (!tables[currentTable]) tables[currentTable] = { name: currentTable, displayName: displayFromCell || displayFallback, values: [] };
      continue;
    }
    if (currentTable && r[2]) {
      tables[currentTable].values.push({ value: String(r[2]).trim(), code: r[3] ? String(r[3]).trim() : null });
    }
  }
  return Object.values(tables);
}

function parseProductTypeIndex(wb, taxName, taxDisplay) {
  const ws = wb.Sheets['Product Type Index'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const nodes = [];
  const l2Seen = {};
  const l3Seen = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[5] || !r[6]) continue;
    const l1Code = r[1] ? String(r[1]).trim() : null;
    const l1Name = r[2] ? String(r[2]).trim() : null;
    const l2Name = r[3] ? String(r[3]).trim() : null;
    const l3Name = r[4] ? String(r[4]).trim() : null;
    const l4Id = String(r[5]).trim();
    const l4Name = String(r[6]).trim();

    const l2Key = `${l1Code}|${l2Name}`;
    const l3Key = `${l1Code}|${l2Name}|${l3Name}`;
    if (!l2Seen[l2Key]) l2Seen[l2Key] = l4Id.substring(0, 4);
    if (!l3Seen[l3Key]) l3Seen[l3Key] = l4Id.substring(0, 7);

    nodes.push({ l1Code, l1Name, l2Name, l3Name, l4Id, l4Name, l2Id: l2Seen[l2Key], l3Id: l3Seen[l3Key] });
  }
  return { name: taxName, displayName: taxDisplay || taxName, nodes };
}

// Resolve a column by header name. Returns -1 when absent.
function colIndex(headerRow, names) {
  const norm = s => String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
  const cells = (headerRow || []).map(norm);
  for (const n of names) {
    const i = cells.indexOf(norm(n));
    if (i >= 0) return i;
  }
  return -1;
}

function parseCategorySpecificAttrs(wb, attrs) {
  const ws = wb.Sheets['Category Specific Attributes'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const attrByDisplay = {};
  for (const a of attrs) attrByDisplay[a.displayName.trim().toLowerCase()] = a;

  // Columns are located by header, so inserting or removing one (e.g. "Default UOM")
  // has no effect on parsing.
  const hdr     = rows[0] || [];
  const cL4Id   = colIndex(hdr, ['L4 Node ID']);
  const cL4Name = colIndex(hdr, ['L4 Product Type']);
  const cAttr   = colIndex(hdr, ['Attribute Display Name']);
  const cKva    = colIndex(hdr, ['KVA / VA']);
  const cLov    = colIndex(hdr, ['LOV Values / Notes', 'LOV Values']);
  const cRef    = colIndex(hdr, ['Ref Table']);

  const missing = [
    ['L4 Node ID', cL4Id], ['Attribute Display Name', cAttr], ['KVA / VA', cKva],
    ['LOV Values / Notes', cLov], ['Ref Table', cRef],
  ].filter(([, i]) => i < 0).map(([n]) => n);
  if (missing.length) {
    throw new Error('Category Specific Attributes is missing column(s): ' + missing.join(', ') +
      '. Header row reads: [' + hdr.filter(Boolean).join(' | ') + ']');
  }

  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[cL4Id] || !r[cAttr]) continue;
    const l4NodeId = String(r[cL4Id]).trim();
    const l4Name = (cL4Name >= 0 && r[cL4Name]) ? String(r[cL4Name]).trim() : '';
    const attrDisplayName = String(r[cAttr]).trim();
    const kvaOrVa = r[cKva] ? String(r[cKva]).trim() : null;
    if (!kvaOrVa) continue;
    const lovValuesRaw = r[cLov] ? String(r[cLov]).trim() : null;
    const refTable = r[cRef] && String(r[cRef]).trim() !== '—' ? String(r[cRef]).trim() : null;

    if (!attrDisplayName || attrDisplayName === '—') continue;
    const matchedAttr = attrByDisplay[attrDisplayName.toLowerCase()] ?? null;
    const isFullLov = lovValuesRaw && lovValuesRaw.toLowerCase().includes('full lov');
    const lovValues = (!lovValuesRaw || isFullLov) ? [] : lovValuesRaw.split(',').map(v => v.trim()).filter(v => v && v !== '—');

    result.push({ l4NodeId, l4Name, attrDisplayName, attrFieldId: matchedAttr?.fieldId ?? null, kvaOrVa, lovValues, fullLov: !!isFullLov, refTable });
  }
  return result;
}

