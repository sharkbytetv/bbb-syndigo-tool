// Generates 080-reference-data.xlsm
// Returns { wb, entityIds } where entityIds[refTableName][value] = uuid
function generateRefData(data) {
  const wb = XLSX.utils.book_new();
  const entityIds = {};

  const allRefNames = data.refTables.map(rt => rt.name);
  const codeAttrCols = allRefNames.map(n => n + 'Code');

  const banner = ['System Attributes', null, null, null, 'Reference Data'];
  const header = ['Action', 'Type', 'ID', 'Name', 'Code'];

  const dataRows = [];

  for (const rt of data.refTables) {
    entityIds[rt.name] = {};
    for (const v of rt.values) {
      const id = generateUUID();
      entityIds[rt.name][v.value] = id;
      dataRows.push(['', rt.name, id, v.value, v.code ?? '']);
    }
  }

  buildAOASheet(wb, 'Entities', [banner, header, ...dataRows]);

  buildAOASheet(wb, 'Relationships', [
    ['System Attributes', null, null, null, null, null, null],
    ['Action', 'Type', 'ID', 'Name', 'Related to ID', 'Related to Type', 'Code'],
  ]);

  return { wb, entityIds };
}
