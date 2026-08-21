// Generates one XLSM per ref table, each containing Entities + Relationships.
// Returns { files: [{ name, displayName, buf }], warnings }.
async function generateRefDataFiles(data) {
  const tableValueToL4s = {};
  const blankWarnSeen = new Set();
  const warnings = [];
  for (const ca of data.categoryAttrs) {
    if (!ca.refTable || !ca.l4NodeId) continue;
    if (ca.fullLov) continue;
    if (ca.lovValues.length === 0) {
      const warnKey = ca.refTable + '|' + ca.attrDisplayName;
      if (!blankWarnSeen.has(warnKey)) {
        blankWarnSeen.add(warnKey);
        warnings.push('Blank LOV in CSA for "' + ca.attrDisplayName + '" (' + ca.refTable + ') — no relationships created for these rows.');
      }
      continue;
    }
    if (!tableValueToL4s[ca.refTable]) tableValueToL4s[ca.refTable] = {};
    const byVal = tableValueToL4s[ca.refTable];
    for (const val of ca.lovValues) {
      const key = val.toLowerCase();
      if (!byVal[key]) byVal[key] = [];
      if (!byVal[key].includes(ca.l4NodeId)) byVal[key].push(ca.l4NodeId);
    }
  }

  const results = [];

  for (const rt of data.refTables) {
    const zip = await loadTemplate('templates/ref-data-template.xlsm');
    const relType = rt.name + '@@Belongs to Product Type';

    const entityRows = [];
    const valueToSeq = {};
    let seq = 1;
    for (const v of rt.values) {
      const key = v.value.toLowerCase();
      valueToSeq[key] = seq;
      // col K ("Value") mirrors the display name, as in the Syndigo export
      entityRows.push(['', rt.name, seq, v.value, v.code ?? '', null, null, null, null, null, v.value]);
      seq++;
    }

    const relRows = [];
    const seen = new Set();
    const byVal = tableValueToL4s[rt.name] || {};
    for (const v of rt.values) {
      const key = v.value.toLowerCase();
      const id = valueToSeq[key];
      for (const l4Code of (byVal[key] || [])) {
        const dedupKey = id + '|' + l4Code;
        if (seen.has(dedupKey)) continue;
        seen.add(dedupKey);
        // "Related to ID" (col E) left null — Syndigo populates it from the Code
        relRows.push(['', relType, id, v.value, null, 'refproducttype', l4Code]);
      }
    }

    // Reference Data has a single header row (Type, ENTITY_NAME); the template
    // ships hardcoded refsizemaster rows, so rewrite it for this table.
    const refDataRows = [[relType, rt.name]];

    // Entities/Relationships have 2 header rows (section header + column headers);
    // their data starts at row 3.
    await fillSheet(zip, 'Entities',       entityRows,  2);
    await fillSheet(zip, 'Relationships',  relRows,     2);
    await fillSheet(zip, 'Reference Data', refDataRows, 1);

    const buf = await zip.generateAsync({ type: 'uint8array' });
    results.push({ name: rt.name, displayName: rt.displayName, buf });
  }

  return { files: results, warnings };
}
