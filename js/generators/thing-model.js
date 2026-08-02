async function generateThingModel(data, tenant) {
  const wb = await loadTemplate('templates/thing-template.xlsx');

  updateMetadata(wb, tenant, 'thing');

  fillSheet(wb, 'ENTITIES', buildThingEntities());
  fillSheet(wb, 'ATTRIBUTES', buildThingAttributes(data));
  fillSheet(wb, 'RELATIONSHIPS', buildThingRelationships());
  fillSheet(wb, 'E-A-R MODEL', buildThingEAR(data));

  return wb;
}

function buildThingEntities() {
  return [
    [null, 'product', 'Product', null, null, null, null],
    [null, 'sku',     'SKU',     null, null, null, null],
  ];
}

function buildThingAttributes(data) {
  const rows = [];
  let seq = 10;
  for (const a of data.attributes) {
    const dt = normalizeDataType(a.dataType);
    const display = inferDisplayType(a);
    const isPath = dt === 'path';
    const row = new Array(59).fill(null);
    row[1]  = a.fieldId;
    row[2]  = 'thing';
    row[3]  = dt;
    row[4]  = 'No';
    row[5]  = a.displayName;
    row[6]  = a.group || null;
    row[7]  = display;
    row[8]  = a.isMultiValued;
    row[9]  = a.isTranslatable;
    row[10] = a.isSyndicatable;
    row[11] = a.isMandatory;
    row[15] = seq;
    row[16] = 'No';
    row[18] = a.description || null;
    row[19] = a.refTable || null;
    row[25] = isPath ? data.taxonomy.name : null;
    row[27] = isPath ? 'Yes' : null;
    row[30] = isPath ? '>' : null;
    row[35] = 'No';
    row[36] = a.isFilterableOnline;
    row[40] = a.minLength ?? null;
    row[41] = a.maxLength ?? null;
    row[42] = a.rangeFrom ?? null;
    row[43] = a.rangeFrom !== null ? a.rangeFromInclusive : null;
    row[44] = a.rangeTo ?? null;
    row[45] = a.rangeTo !== null ? a.rangeToInclusive : null;
    row[46] = a.precision ?? null;
    row[47] = a.minPrecision ?? null;
    row[48] = a.maxPrecision ?? null;
    if (a.dependentAttribute) {
      row[49] = a.dependentAttribute;
    }
    rows.push(row);
    seq += 10;
  }
  return rows;
}

function buildThingRelationships() {
  return [
    [null, 'productToSku', 'thing', 'Composition', 'SKUs', 'Product', null, null, null, null, null, null, null],
  ];
}

function buildThingEAR(data) {
  const rows = [];

  const isId = (entity, a) => {
    if (entity === 'product' && a.fieldId === 'thgproductid') return 'Yes';
    if (entity === 'sku'     && a.fieldId === 'thgskuid')     return 'Yes';
    return 'No';
  };
  const isExtName = (entity, a) => {
    if (entity === 'product' && a.fieldId === 'thgproductname') return 'Yes';
    return 'No';
  };

  const makeRow = (entity, a) => {
    const row = new Array(50).fill(null);
    row[1]  = entity;
    row[3]  = a.fieldId;
    row[8]  = isId(entity, a);
    row[9]  = isExtName(entity, a);
    row[10] = a.group || null;
    row[13] = 'No';
    return row;
  };

  for (const a of data.attributes) {
    const onProduct = a.isProductLevel === 'Yes';
    const onSku     = a.isSkuVariant   === 'Yes';
    if (onProduct || (!onProduct && !onSku)) rows.push(makeRow('product', a));
    if (onSku)                               rows.push(makeRow('sku', a));
  }

  const relRow = new Array(50).fill(null);
  relRow[1] = 'product';
  relRow[2] = 'productToSku';
  relRow[5] = 'sku';
  rows.push(relRow);

  return rows;
}
