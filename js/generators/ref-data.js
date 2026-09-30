// Generates the reference-data files, each containing Entities + Relationships.
// `combine` emits a single workbook covering every table instead of one file per table.
// Returns { files: [{ name, displayName, buf, tables }], warnings }.
async function generateRefDataFiles(data, combine) {
  // Build map from CSA specific-LOV rows: refTable → { valueLower → [l4Code, ...] }
  // fullLov rows (Syndigo implicit) and blank rows (warned below) do not create relationships.
  const tableValueToL4s = {};
  const blankWarnSeen = new Set();
  const warnings = [];
  for (const ca of data.categoryAttrs) {
    if (!ca.refTable || !ca.l4NodeId) continue;
    if (ca.fullLov) continue; // Syndigo applies full LOV implicitly — no relationships needed
    if (ca.lovValues.length === 0) {
      const warnKey = ca.refTable + '|' + ca.attrDisplayName;
      if (!blankWarnSeen.has(warnKey)) {
        blankWarnSeen.add(warnKey);
        warnings.push('Blank LOV in CSA for "' + ca.attrDisplayName + '" (' + ca.refTable +
          ') — treated as full LOV, so no relationships are created. Write "Full LOV for Ref Table" to make that explicit.');
      }
      continue;
    }
    // Keyed lower-case: the CSA tab and the Reference LOV Master drift in casing
    // (refSize vs refsize), and an exact match silently yields no relationships.
    const tableKey = matchKey(ca.refTable);
    if (!tableValueToL4s[tableKey]) tableValueToL4s[tableKey] = {};
    const byVal = tableValueToL4s[tableKey];
    for (const val of ca.lovValues) {
      const key = matchKey(val);
      if (!byVal[key]) byVal[key] = [];
      if (!byVal[key].includes(ca.l4NodeId)) byVal[key].push(ca.l4NodeId);
    }
  }

  const groups = [];

  // Placeholder ids only have to be unique within the file they land in. Split output
  // restarts at 1 per table; combined output keeps counting, so two tables' new values
  // can never share an id in the same sheet.
  let seq = 1;

  for (const rt of data.refTables) {
    if (!combine) seq = 1;
    const relType = rt.name + '@@Belongs to Product Type';

    // Entities: assign sequence IDs
    const entityRows = [];
    const valueToSeq = {};
    for (const v of rt.values) {
      const key = matchKey(v.value);
      valueToSeq[key] = seq;
      entityRows.push(['', rt.name, seq, v.value, v.code ?? '', null, null, null, null, null, v.value]);
      seq++;
    }

    // Relationships: one row per (entity, l4Code) pair — specific LOV rows only
    const relRows = [];
    const seen = new Set();
    const byVal = tableValueToL4s[matchKey(rt.name)] || {};
    for (const v of rt.values) {
      const key = matchKey(v.value);
      const id = valueToSeq[key];
      for (const l4Code of (byVal[key] || [])) {
        const dedupKey = id + '|' + l4Code;
        if (seen.has(dedupKey)) continue;
        seen.add(dedupKey);
        relRows.push(['', relType, id, v.value, null, 'refproducttype', l4Code]);
      }
    }

    groups.push({
      name: rt.name, displayName: rt.displayName, typeName: rt.name,
      relType, entityRows, relRows,
    });
  }

  const files = await writeRefDataWorkbooks(groups, combine);
  return { files, warnings };
}
