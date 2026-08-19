async function generateTaxonomyModel(data, tenant) {
  const zip = await loadTemplate('templates/tax-template.xlsx');

  await updateMetadata(zip, tenant, 'taxonomyModel');
  await fillSheet(zip, 'TAXONOMIES',    buildTaxonomies(data.taxonomy));
  await fillSheet(zip, 'CATEGORIES',    buildCategories(data.taxonomy));
  await fillSheet(zip, 'CATEGORY MODEL', []);  // clear template placeholder rows; not populated

  return zip.generateAsync({ type: 'uint8array' });
}

function buildTaxonomies(taxonomy) {
  return [[null, taxonomy.name, taxonomy.displayName, null]];
}

function buildCategories(taxonomy) {
  const rows = [];
  const taxName = taxonomy.name;
  const l1Added = new Set();
  const l2Added = new Set();
  const l3Added = new Set();

  for (const n of taxonomy.nodes) {
    if (!l1Added.has(n.l1Code)) {
      rows.push([null, n.l1Code, 'Classification', null,      taxName, n.l1Name, null]);
      l1Added.add(n.l1Code);
    }
    if (!l2Added.has(n.l2Id)) {
      rows.push([null, n.l2Id, 'Classification', n.l1Code,   taxName, n.l2Name, null]);
      l2Added.add(n.l2Id);
    }
    if (!l3Added.has(n.l3Id)) {
      rows.push([null, n.l3Id, 'Classification', n.l2Id,     taxName, n.l3Name, null]);
      l3Added.add(n.l3Id);
    }
    rows.push([null, n.l4Id, 'Classification', n.l3Id, taxName, n.l4Name, null]);
  }
  return rows;
}

function buildCategoryModel(data) {
  const rows = [];
  const seen = new Set();
  const taxName = data.taxonomy.name;

  for (const ca of data.categoryAttrs) {
    if (!ca.attrFieldId) continue;
    const key = ca.l4NodeId + '|' + ca.attrFieldId;
    if (seen.has(key)) continue;
    seen.add(key);

    const attr = data.attributes.find(a => a.fieldId === ca.attrFieldId);
    const row = new Array(44).fill(null);
    row[1]  = taxName;
    row[2]  = ca.l4NodeId;
    row[3]  = 'Classification';
    row[4]  = ca.attrFieldId;
    row[6]  = attr?.group || null;
    row[17] = ca.kvaOrVa === 'KVA' ? 'Yes' : 'No';
    row[18] = attr?.isMultiValued  || 'No';
    row[19] = attr?.isTranslatable || 'No';
    row[20] = attr?.isSyndicatable || 'No';
    row[22] = ca.refTable || attr?.refTable || null;
    rows.push(row);
  }
  return rows;
}
