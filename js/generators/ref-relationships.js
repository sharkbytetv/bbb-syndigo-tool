// Generates 100-reference-data-relationship.xlsm
// entityIds: { refTableName: { value: uuid } } from ref-data generator
// NOTE: Related to ID uses the L4 Node ID as a placeholder.
// After the taxonomy model is imported into Syndigo, replace these with the
// actual internal entity GUIDs for each taxonomy category node.
function generateRefRelationships(data, entityIds) {
  const wb = XLSX.utils.book_new();

  const banner = ['System Attributes', null, null, null, null, null, null];
  const header = ['Action', 'Type', 'ID', 'Name', 'Related to ID', 'Related to Type', 'Code'];

  const dataRows = [];
  const seen = new Set();

  for (const ca of data.categoryAttrs) {
    if (!ca.refTable || ca.lovValues.length === 0) continue;
    const tableIds = entityIds[ca.refTable];
    if (!tableIds) continue;

    const relType = ca.refTable + 'ToCategory';

    for (const val of ca.lovValues) {
      const matchedKey = Object.keys(tableIds).find(k => k.toLowerCase() === val.toLowerCase());
      if (!matchedKey) continue;

      const entityId = tableIds[matchedKey];
      const dedupKey = `${entityId}|${ca.l4NodeId}`;
      if (seen.has(dedupKey)) continue;
      seen.add(dedupKey);

      dataRows.push([
        '',
        relType,
        entityId,
        matchedKey,
        ca.l4NodeId,
        'category',
        '',
      ]);
    }
  }

  buildAOASheet(wb, 'Relationships', [banner, header, ...dataRows]);
  buildAOASheet(wb, 'Entities', [
    ['System Attributes', null, null, null],
    ['Action', 'Type', 'ID', 'Name'],
  ]);

  return wb;
}
