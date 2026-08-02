async function generateRefModel(data, tenant) {
  const wb = await loadTemplate('templates/ref-template.xlsx');
  const refTablesWithCSA = getRefTablesWithCSA(data);

  updateMetadata(wb, tenant, 'referenceData');
  fillSheet(wb, 'ENTITIES',      buildRefEntities(data.refTables));
  fillSheet(wb, 'ATTRIBUTES',    buildRefAttributes(data.refTables));
  fillSheet(wb, 'RELATIONSHIPS', buildRefRelationships(refTablesWithCSA));
  fillSheet(wb, 'E-A-R MODEL',   buildRefEAR(data.refTables, refTablesWithCSA));

  return wb;
}

function getRefTablesWithCSA(data) {
  const s = new Set();
  for (const ca of data.categoryAttrs) if (ca.refTable) s.add(ca.refTable);
  return s;
}

function buildRefEntities(refTables) {
  // 7 cols: ACTION, NAME, DISPLAY NAME, ICON, MERGE SEQUENCE, HELP TEXT, BASE UNIT SYMBOL
  return refTables.map(rt => [null, rt.name, rt.displayName, null, null, null, null]);
}

function buildRefAttributes(refTables) {
  // 59 cols — same template as thing domain
  const rows = [];
  let seq = 10;
  for (const rt of refTables) {
    const row = new Array(59).fill(null);
    row[1]  = rt.name + 'Code';
    row[2]  = 'referenceData';
    row[3]  = 'string';
    row[4]  = 'No';
    row[5]  = 'Code';
    row[7]  = 'textbox';
    row[8]  = 'No';
    row[9]  = 'No';
    row[10] = 'No';
    row[11] = 'No';
    row[15] = seq;
    row[16] = 'No';
    row[18] = 'Short code identifier';
    row[35] = 'No';
    rows.push(row);
    seq += 10;
  }
  return rows;
}

function buildRefRelationships(refTablesWithCSA) {
  // 13 cols: ACTION, NAME, DOMAIN, RELATIONSHIP TYPE, DISPLAY NAME, DISPLAY NAME WHEREUSED,
  //          DISPLAY SEQUENCE, RELATED ENTITY SEARCH, HELP TEXT, MIN OCCURRENCE, MAX OCCURRENCE,
  //          IGNORE MERGE?, MERGE SEQUENCE
  const rows = [];
  for (const refName of refTablesWithCSA) {
    const row = new Array(13).fill(null);
    row[1] = refName + 'ToCategory';
    row[2] = 'referenceData';
    row[3] = 'Association';
    row[4] = 'Product Types';
    row[5] = 'Applicable Values';
    rows.push(row);
  }
  return rows;
}

function buildRefEAR(refTables, refTablesWithCSA) {
  // 50 cols — same template as thing domain
  const rows = [];
  for (const rt of refTables) {
    // Code attribute row — entity identifier
    const attrRow = new Array(50).fill(null);
    attrRow[1]  = rt.name;
    attrRow[3]  = rt.name + 'Code';
    attrRow[8]  = 'Yes';   // IS ENTITY IDENTIFIER?
    attrRow[9]  = 'No';
    attrRow[13] = 'No';
    rows.push(attrRow);

    // Relationship row → taxonomy category (if in CSA)
    if (refTablesWithCSA.has(rt.name)) {
      const relRow = new Array(50).fill(null);
      relRow[1] = rt.name;
      relRow[2] = rt.name + 'ToCategory';
      relRow[5] = 'category';
      rows.push(relRow);
    }
  }
  return rows;
}
