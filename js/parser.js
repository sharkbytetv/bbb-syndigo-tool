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
      currentTable = cell0.split('—')[0].trim();
      if (!tables[currentTable]) tables[currentTable] = { name: currentTable, displayName: r[1] ? String(r[1]).trim() : currentTable, values: [] };
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

function parseCategorySpecificAttrs(wb, attrs) {
  const ws = wb.Sheets['Category Specific Attributes'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const attrByDisplay = {};
  for (const a of attrs) attrByDisplay[a.displayName.trim().toLowerCase()] = a;

  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[4] || !r[6]) continue;
    const l4NodeId = String(r[4]).trim();
    const l4Name = r[5] ? String(r[5]).trim() : '';
    const attrDisplayName = String(r[6]).trim();
    const kvaOrVa = r[7] ? String(r[7]).trim() : null;
    if (!kvaOrVa) continue;
    const lovValuesRaw = r[8] ? String(r[8]).trim() : null;
    const refTable = r[9] && String(r[9]).trim() !== '—' ? String(r[9]).trim() : null;

    if (!attrDisplayName || attrDisplayName === '—') continue;
    const matchedAttr = attrByDisplay[attrDisplayName.toLowerCase()] ?? null;
    const lovValues = lovValuesRaw ? lovValuesRaw.split(',').map(v => v.trim()).filter(v => v && v !== '—') : [];

    result.push({ l4NodeId, l4Name, attrDisplayName, attrFieldId: matchedAttr?.fieldId ?? null, kvaOrVa, lovValues, refTable });
  }
  return result;
}
