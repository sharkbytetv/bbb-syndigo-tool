// Delta mode for reference data.
//
// Compares a combined current-model export (all ref tables in one workbook) against the DD
// and emits, per ref table, only the rows needed to bring the environment in line:
//   • values in the DD but not the model      → add row (blank Action)
//   • values in the model but not the DD      → DELETE row (carries the model's UUID)
//   • relationships                           → full diff, see buildRelationshipDelta below
//
// Values are matched on Code. The DD is expected to always populate it; anything missing a
// code is skipped in both directions and reported, so a blank code can never cause a delete.

const REL_SUFFIX       = '@@Belongs to Product Type';
const DELETE_ACTION    = 'DELETE';
const NEW_REL_REL_TYPE = 'refproducttype';

// Locate the header row (exports carry a section row above it) and map column name → index.
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// Some exports come out with no header rows at all, and sometimes without the leading Action
// column. Column order is fixed either way, so anchor on the ID column — found by looking for
// UUID-shaped values — and derive the rest from it.
//   Entities:      … Type | ID | Name | Code …
//   Relationships: … Type | ID | Name | Related to ID | Related to Type | Code …
function headerlessIndex(rows, kind) {
  const votes = {};
  for (let r = 0; r < Math.min(25, rows.length); r++) {
    (rows[r] || []).forEach((c, i) => {
      if (UUID_RE.test(String(c ?? '').trim())) votes[i] = (votes[i] || 0) + 1;
    });
  }
  const idIdx = Object.keys(votes).sort((a, b) => votes[b] - votes[a])[0];
  if (idIdx === undefined) return null;
  const id = Number(idIdx);
  if (id < 1) return null;                       // Type must sit to the left of ID

  const map = { 'Type': id - 1, 'ID': id, 'Name': id + 1, _dataStart: 0, _headerless: true };
  if (kind === 'entities') {
    map['Code'] = id + 2;
  } else {
    map['Related to ID']   = id + 2;
    map['Related to Type'] = id + 3;
    map['Code']            = id + 4;
  }
  return map;
}

function headerIndex(rows, required) {
  for (let r = 0; r < Math.min(8, rows.length); r++) {
    const cells = (rows[r] || []).map(c => (c === null || c === undefined) ? '' : String(c).trim());
    const map = {};
    cells.forEach((c, i) => { if (c && map[c] === undefined) map[c] = i; });
    if (required.every(k => map[k] !== undefined)) {
      map._dataStart = r + 1;
      return map;
    }
  }
  return null;
}

// Headers win when present; only fall back to positional inference when they are absent,
// otherwise the header rows themselves get read as data.
function resolveSheet(rows, required, kind, sheetName, warnings) {
  const byHeader = headerIndex(rows, required);
  if (byHeader) return byHeader;

  const positional = headerlessIndex(rows, kind);
  if (positional) {
    warnings.push('"' + sheetName + '" sheet has no header row — columns inferred from the ' +
      'position of the ID column (Type=' + positional['Type'] + ', ID=' + positional['ID'] +
      ', Name=' + positional['Name'] + ', Code=' + positional['Code'] + ').');
    return positional;
  }

  const seen = [];
  for (let r = 0; r < Math.min(4, rows.length); r++) {
    const cells = (rows[r] || []).map(c => (c === null || c === undefined) ? '' : String(c).trim());
    seen.push('row ' + (r + 1) + ': [' + cells.filter(Boolean).slice(0, 8).join(' | ') + ']');
  }
  throw new Error('"' + sheetName + '" — no header row with: ' + required.join(', ') +
    ', and no UUID column to infer from. ' + rows.length + ' row(s). ' +
    (seen.length ? seen.join('  ') : '(sheet is empty)'));
}

const norm = v => (v === null || v === undefined) ? '' : String(v).trim();
const lower = v => norm(v).toLowerCase();

// Parse the combined export into per-table entity and relationship maps.
function parseCurrentModel(wb) {
  const warnings = [];
  const entitiesByTable = {};   // lower(tableName) → entities
  const relsByTable = {};       // lower(tableName) → relationship rows
  const tableCasing = {};       // lower(tableName) → the model's own spelling

  const eSheet = wb.Sheets['Entities'];
  if (!eSheet) {
    throw new Error('No "Entities" sheet. Sheets in this file: ' + (wb.SheetNames || []).join(', '));
  }
  const eRows = XLSX.utils.sheet_to_json(eSheet, { header: 1, defval: null, blankrows: true });
  const eh = resolveSheet(eRows, ['Type', 'ID', 'Name', 'Code'], 'entities', 'Entities', warnings);

  for (let i = eh._dataStart; i < eRows.length; i++) {
    const row = eRows[i] || [];
    const table = norm(row[eh['Type']]);
    if (!table) continue;
    const entity = {
      uuid: norm(row[eh['ID']]),
      name: norm(row[eh['Name']]),
      code: norm(row[eh['Code']]),
    };
    if (!entity.name && !entity.code) continue;
    // Table names are keyed lower-case: the DD writes refSize where the model holds refsize
    const key = lower(table);
    if (!entitiesByTable[key]) { entitiesByTable[key] = []; tableCasing[key] = table; }
    entitiesByTable[key].push(entity);
  }

  const rSheet = wb.Sheets['Relationships'];
  if (rSheet) {
    const rRows = XLSX.utils.sheet_to_json(rSheet, { header: 1, defval: null, blankrows: true });
    const rh = resolveSheet(rRows, ['Type', 'ID', 'Name', 'Related to ID'], 'relationships',
                            'Relationships', warnings);
    for (let i = rh._dataStart; i < rRows.length; i++) {
      const row = rRows[i] || [];
      const relType = norm(row[rh['Type']]);
      if (!relType) continue;
      const table = lower(relType.split('@@')[0]);
      if (!table) continue;
      if (!relsByTable[table]) relsByTable[table] = [];
      relsByTable[table].push({
        relType:       relType,
        uuid:          norm(row[rh['ID']]),
        name:          norm(row[rh['Name']]),
        relatedToId:   norm(row[rh['Related to ID']]),
        relatedToType: rh['Related to Type'] !== undefined ? norm(row[rh['Related to Type']]) : '',
        // Code carries the L4 node id — this is the join key against the DD
        code:          rh['Code'] !== undefined ? norm(row[rh['Code']]) : '',
      });
    }
  } else {
    warnings.push('Current model export has no "Relationships" sheet — relationship diffing skipped.');
  }

  return { entitiesByTable, relsByTable, tableCasing, warnings };
}

// Per ref table, work out which L4 nodes the DD constrains and which it opens up.
//   specificNodes → node has an explicit LOV list; relationships should mirror it exactly
//   fullLovNodes  → node is "full LOV"; ALL existing relationships must go, since the
//                   absence of relationships is what signals "unrestricted" to Syndigo
//   desiredRels   → set of "valueLower|l4Code" the DD wants to exist
function buildDdRelationshipView(categoryAttrs, tableName) {
  const specificNodes = new Set();
  const fullLovNodes  = new Set();
  const blankNodes    = new Set();
  const desiredRels   = new Set();

  const want = lower(tableName);
  for (const ca of categoryAttrs) {
    if (lower(ca.refTable) !== want || !ca.l4NodeId) continue;
    if (ca.fullLov) { fullLovNodes.add(ca.l4NodeId); continue; }
    if (!ca.lovValues || ca.lovValues.length === 0) { blankNodes.add(ca.l4NodeId); continue; }
    specificNodes.add(ca.l4NodeId);
    for (const v of ca.lovValues) desiredRels.add(lower(v) + '|' + ca.l4NodeId);
  }
  // A node listed as full LOV anywhere wins — it is unrestricted regardless of other rows.
  for (const n of fullLovNodes) specificNodes.delete(n);
  return { specificNodes, fullLovNodes, blankNodes, desiredRels };
}

// The L4 node a relationship targets lives in the Code column, matching what the full-import
// generator writes (Related to Type = refproducttype, Related to ID blank, Code = l4NodeId).
// Returns null when Code is blank or names a node the DD does not know about; those rows are
// left alone rather than deleted.
function resolveL4(rel, validL4Ids) {
  return (rel.code && validL4Ids.has(rel.code)) ? rel.code : null;
}

function buildRelationshipDelta(opts) {
  const { tableName, modelRels, ddView, deletedNames, liveValueIds, validL4Ids, nameMap, warnings } = opts;
  const relType = tableName + REL_SUFFIX;
  const addRows = [];
  const delRows = [];

  // Translate the stored value name to the DD's name before comparing
  const canon = n => (nameMap && nameMap[lower(n)]) || lower(n);

  const modelRelSet = new Set();
  let unresolved = 0;
  for (const r of modelRels) {
    r._l4 = resolveL4(r, validL4Ids);
    if (r._l4) modelRelSet.add(canon(r.name) + '|' + r._l4);
    else unresolved++;
  }
  if (unresolved) {
    warnings.push('[' + tableName + '] ' + unresolved + ' existing relationship row(s) have a blank Code, ' +
      'or a Code naming an L4 node the DD does not contain — left untouched (never deleted).');
  }

  // Deletions — driven by what the model currently holds
  for (const r of modelRels) {
    let drop = false;
    if (deletedNames.has(lower(r.name))) {
      drop = true;                                        // value itself is going away
    } else if (r._l4 && ddView.fullLovNodes.has(r._l4)) {
      drop = true;                                        // node became unrestricted
    } else if (r._l4 && ddView.specificNodes.has(r._l4) &&
               !ddView.desiredRels.has(canon(r.name) + '|' + r._l4)) {
      drop = true;                                        // value no longer listed for this node
    }
    // Unresolvable rows, and nodes the DD says nothing about, are left untouched.
    if (drop) {
      delRows.push([DELETE_ACTION, r.relType || relType, r.uuid, r.name, r.relatedToId || null,
                    r.relatedToType || NEW_REL_REL_TYPE, r.code || null, null]);
    }
  }

  // Additions — anything the DD wants that the model does not already have
  for (const key of ddView.desiredRels) {
    if (modelRelSet.has(key)) continue;
    const sep     = key.lastIndexOf('|');
    const valLower = key.slice(0, sep);
    const l4Code   = key.slice(sep + 1);
    const live     = liveValueIds[valLower];
    if (!live) {
      warnings.push('[' + tableName + '] CSA lists "' + valLower + '" for node ' + l4Code +
                    ' but that value is not in the reference table — relationship skipped.');
      continue;
    }
    // Same shape as the full-import generator: Related to ID blank, L4 node in Code
    addRows.push(['', relType, live.id, live.name, null, NEW_REL_REL_TYPE, l4Code, null]);
  }

  return { addRows, delRows };
}

// Generates one XLSM per CHANGED ref table. Returns { files, warnings, summary }.
async function generateRefDataDelta(data, modelWb) {
  const model = parseCurrentModel(modelWb);
  const warnings = model.warnings.slice();
  const results = [];
  const summary = [];

  // The DD's L4 node ID space — used to resolve which node each existing relationship targets
  const validL4Ids = new Set();
  for (const n of (data.taxonomy?.nodes || [])) if (n.l4Id) validL4Ids.add(norm(n.l4Id));
  for (const ca of data.categoryAttrs) if (ca.l4NodeId) validL4Ids.add(norm(ca.l4NodeId));
  if (!validL4Ids.size) {
    warnings.push('No L4 node IDs found in the DD — relationship diffing will be skipped entirely.');
  }

  for (const rt of data.refTables) {
    const key = lower(rt.name);
    const modelEntities = model.entitiesByTable[key] || [];
    const modelRels     = model.relsByTable[key]     || [];
    // Emit rows under the model's own spelling — that is the entity type Syndigo knows
    const typeName = model.tableCasing[key] || rt.name;

    if (!model.entitiesByTable[key]) {
      warnings.push('[' + rt.name + '] not present in the current model export — treating every DD value as new.');
    }

    // Index the model by code, keeping name indexes for values whose Code is missing
    const modelByCode = {}, modelByName = {}, modelBlankByName = {};
    for (const e of modelEntities) {
      if (e.code) modelByCode[lower(e.code)] = e;
      else        modelBlankByName[lower(e.name)] = e;
      modelByName[lower(e.name)] = e;
    }

    // Index the DD by code, keeping a name index for values whose Code is missing
    const ddByCode = {}, ddByName = {}, ddBlankByName = {};
    for (const v of rt.values) {
      if (v.code) ddByCode[lower(v.code)] = v;
      else {
        ddBlankByName[lower(v.value)] = v;
        warnings.push('[' + rt.name + '] DD value "' + v.value +
          '" has no Code in Reference LOV Master — matched by name instead.');
      }
      ddByName[lower(v.value)] = v;
    }

    // Pair model entities to DD values. Code is the match key; the name fallback applies only
    // when one side has no Code — without it, a model value missing its Code would be deleted
    // and re-added, losing its UUID and every relationship hanging off it.
    const pairs = [];
    const matchedModel = new Set();
    const matchedDd = new Set();
    for (const v of rt.values) {
      const e = v.code
        ? (modelByCode[lower(v.code)] || modelBlankByName[lower(v.value)])
        : modelByName[lower(v.value)];
      if (e && !matchedModel.has(e)) {
        pairs.push({ entity: e, value: v });
        matchedModel.add(e);
        matchedDd.add(v);
      }
    }

    // Relationship rows carry the value's name as stored in the model, which may be stale.
    // Map it to the DD's name so those rows still line up with the CSA's LOV entries.
    const nameMap = {};
    for (const p of pairs) nameMap[lower(p.entity.name)] = lower(p.value.value);

    // Value-level diff
    const addRows = [];
    const updRows = [];
    const delRows = [];
    const deletedNames = new Set();
    const liveValueIds = {};   // valueLower → { id, name } for relationship rows

    // Matched values keep their UUID; re-state them when Name or Code drifted from the DD
    for (const p of pairs) {
      liveValueIds[lower(p.value.value)] = { id: p.entity.uuid, name: p.value.value };
      const nameDrift = norm(p.entity.name) !== norm(p.value.value);
      const codeDrift = norm(p.entity.code) !== norm(p.value.code || '');
      if (nameDrift || codeDrift) {
        updRows.push(['', typeName, p.entity.uuid, p.value.value, p.value.code || null,
                      null, null, null, null, null, p.value.value]);
      }
    }

    // New values get sequence placeholders, matching the full-import generator's convention
    let seq = 1;
    for (const v of rt.values) {
      if (matchedDd.has(v)) continue;
      addRows.push(['', typeName, seq, v.value, v.code || null, null, null, null, null, null, v.value]);
      liveValueIds[lower(v.value)] = { id: seq, name: v.value };
      seq++;
    }

    for (const e of modelEntities) {
      if (matchedModel.has(e)) continue;
      deletedNames.add(lower(e.name));
      delRows.push([DELETE_ACTION, typeName, e.uuid, e.name, e.code || null, null, null, null, null, null, e.name]);
      if (!e.code) {
        warnings.push('[' + rt.name + '] model value "' + e.name +
          '" has no Code and no matching name in the DD — marked for DELETE.');
      }
    }

    // Relationship diff
    const ddView = buildDdRelationshipView(data.categoryAttrs, rt.name);
    const rel = buildRelationshipDelta({
      tableName: typeName, modelRels, ddView, deletedNames, liveValueIds, validL4Ids, nameMap, warnings,
    });

    const entityRows = addRows.concat(updRows, delRows);
    const relRows    = rel.addRows.concat(rel.delRows);

    if (!entityRows.length && !relRows.length) continue;   // nothing to do for this table

    const zip = await loadTemplate('templates/ref-data-template.xlsm');
    await fillSheet(zip, 'Entities',       entityRows, 2);
    await fillSheet(zip, 'Relationships',  relRows,    2);
    await fillSheet(zip, 'Reference Data', [[typeName + REL_SUFFIX, typeName]], 1);

    const buf = await zip.generateAsync({ type: 'uint8array' });
    results.push({ name: rt.name, displayName: rt.displayName, buf });
    summary.push({
      table:      typeName,
      added:      addRows.length,
      updated:    updRows.length,
      deleted:    delRows.length,
      relAdded:   rel.addRows.length,
      relDeleted: rel.delRows.length,
    });
  }

  return { files: results, warnings, summary };
}

