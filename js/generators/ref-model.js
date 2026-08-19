async function generateRefModel(data, tenant) {
  const zip = await loadTemplate('templates/ref-template.xlsx');
  const refTablesWithCSA = getRefTablesWithCSA(data);

  await updateMetadata(zip, tenant, 'referenceData');
  await fillSheet(zip, 'ENTITIES',      buildRefEntities(data.refTables));
  await fillSheet(zip, 'ATTRIBUTES',    buildRefAttributes(data.refTables));
  await fillSheet(zip, 'RELATIONSHIPS', buildRefRelationships(refTablesWithCSA));
  await fillSheet(zip, 'E-A-R MODEL',   buildRefEAR(data.refTables, refTablesWithCSA));

  return zip.generateAsync({ type: 'uint8array' });
}

function getRefTablesWithCSA(data) {
  const s = new Set();
  for (const ca of data.categoryAttrs) if (ca.refTable) s.add(ca.refTable);
  return s;
}

function buildRefEntities(refTables) {
  return refTables.map(rt => [null, rt.name, rt.displayName, null, null, null, null]);
}

function buildRefAttributes(refTables) {
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
  const rows = [];
  for (const rt of refTables) {
    const attrRow = new Array(50).fill(null);
    attrRow[1]  = rt.name;
    attrRow[3]  = rt.name + 'Code';
    attrRow[8]  = 'Yes';
    attrRow[9]  = 'No';
    attrRow[13] = 'No';
    rows.push(attrRow);

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
