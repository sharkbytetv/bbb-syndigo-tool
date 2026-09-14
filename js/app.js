let rawWorkbook = null;
let modelWorkbook = null;

document.addEventListener('DOMContentLoaded', () => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const deltaToggle = document.getElementById('deltaToggle');
  const deltaPanel = document.getElementById('deltaPanel');
  const modelDropzone = document.getElementById('modelDropzone');
  const modelFileInput = document.getElementById('modelFileInput');
  const modelSummaryEl = document.getElementById('modelSummary');
  const warningsEl = document.getElementById('warnings');
  const deltaSummaryEl = document.getElementById('deltaSummary');
  const tenantInput = document.getElementById('tenantInput');
  const taxNameInput = document.getElementById('taxNameInput');
  const taxDisplayInput = document.getElementById('taxDisplayInput');
  const generateBtn = document.getElementById('generateBtn');
  const statusEl = document.getElementById('status');
  const summaryEl = document.getElementById('summary');

  const badge = document.getElementById('version-badge');
  if (badge && typeof APP_VERSION !== 'undefined') {
    badge.textContent = 'v' + APP_VERSION + ' · ' + APP_BUILT;
    document.title = 'Syndigo Model Generator v' + APP_VERSION;
  }

  // Restore cached values
  ['tenantInput', 'taxNameInput', 'taxDisplayInput'].forEach(id => {
    const saved = localStorage.getItem('syndigo_' + id);
    if (saved) document.getElementById(id).value = saved;
  });

  // Save on change
  [tenantInput, taxNameInput, taxDisplayInput].forEach(el => {
    el.addEventListener('input', () => localStorage.setItem('syndigo_' + el.id, el.value));
  });

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag-over'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); });
  generateBtn.addEventListener('click', generate);

  // Delta mode — restore preference, toggle the panel
  deltaToggle.checked = localStorage.getItem('syndigo_deltaMode') === '1';
  deltaPanel.style.display = deltaToggle.checked ? 'block' : 'none';
  deltaToggle.addEventListener('change', () => {
    localStorage.setItem('syndigo_deltaMode', deltaToggle.checked ? '1' : '0');
    deltaPanel.style.display = deltaToggle.checked ? 'block' : 'none';
  });

  modelDropzone.addEventListener('click', () => modelFileInput.click());
  modelDropzone.addEventListener('dragover', e => { e.preventDefault(); modelDropzone.classList.add('drag-over'); });
  modelDropzone.addEventListener('dragleave', () => modelDropzone.classList.remove('drag-over'));
  modelDropzone.addEventListener('drop', e => {
    e.preventDefault();
    modelDropzone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleModelFile(e.dataTransfer.files[0]);
  });
  modelFileInput.addEventListener('change', () => { if (modelFileInput.files[0]) handleModelFile(modelFileInput.files[0]); });

  function handleModelFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xlsm')) {
      setStatus('Please upload an .xlsx or .xlsm file.', 'error');
      return;
    }
    setStatus('Reading current model export…', 'info');
    const reader = new FileReader();
    reader.onload = e => {
      try {
        modelWorkbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const peek = parseCurrentModel(modelWorkbook);
        const tables = Object.keys(peek.entitiesByTable);
        const valueCount = tables.reduce((n, t) => n + peek.entitiesByTable[t].length, 0);
        const relCount = Object.keys(peek.relsByTable)
          .reduce((n, t) => n + peek.relsByTable[t].length, 0);
        modelSummaryEl.innerHTML = `
          <strong>${file.name}</strong> loaded as current model.<br>
          <span class="pill">${tables.length} reference tables</span>
          <span class="pill">${valueCount} existing values</span>
          <span class="pill">${relCount} existing relationships</span>
        `;
        modelSummaryEl.style.display = 'block';
        setStatus('', '');
      } catch (err) {
        modelWorkbook = null;
        modelSummaryEl.style.display = 'none';
        setStatus('Error reading model export: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function setStatus(msg, type = 'info') {
    statusEl.textContent = msg;
    statusEl.className = 'status ' + type;
    statusEl.style.display = msg ? 'block' : 'none';
  }

  function showDeltaSummary(summary, totals) {
    if (!summary || !summary.length) { deltaSummaryEl.style.display = 'none'; return; }
    const cell = (v, cls) => '<td class="n' + (v ? (cls ? ' ' + cls : '') : ' zero') + '">' + v + '</td>';
    const weight = s => s.added + s.updated + s.deleted + s.relAdded + s.relDeleted;
    const rows = summary
      .slice()
      .sort((a, b) => weight(b) - weight(a))
      .map(s => '<tr><td>' + s.table + '</td>' + cell(s.added) + cell(s.updated) +
                cell(s.deleted, 'del') + cell(s.relAdded) + cell(s.relDeleted, 'del') +
                cell(s.carried) + '</tr>')
      .join('');
    deltaSummaryEl.innerHTML =
      '<table><thead><tr><th>Reference table</th><th>Added</th><th>Updated</th>' +
      '<th>Deleted</th><th>Rel +</th><th>Rel &minus;</th>' +
      '<th title="Unchanged values re-stated so their relationship rows resolve">Carried</th>' +
      '</tr></thead><tbody>' + rows +
      '</tbody><tfoot><tr><td>' + summary.length + ' table(s)</td>' +
      cell(totals.added) + cell(totals.updated) + cell(totals.deleted, 'del') +
      cell(totals.relAdded) + cell(totals.relDeleted, 'del') + cell(totals.carried) +
      '</tr></tfoot></table>';
    deltaSummaryEl.style.display = 'block';
  }

  function showWarnings(list) {
    if (!list || !list.length) { warningsEl.style.display = 'none'; return; }
    // Blank codes mean a value was skipped entirely — call those out first and in red
    const isBlankCode = w => w.includes('has no Code');
    const ordered = list.slice().sort((a, b) => (isBlankCode(b) ? 1 : 0) - (isBlankCode(a) ? 1 : 0));
    const items = ordered.map(w =>
      '<li' + (isBlankCode(w) ? ' class="critical"' : '') + '>' +
      w.replace(/[<>&]/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c])) + '</li>').join('');
    warningsEl.innerHTML = '<h4>' + list.length + ' warning(s)</h4><ul>' + items + '</ul>';
    warningsEl.style.display = 'block';
  }

  function handleFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xlsm')) {
      setStatus('Please upload an .xlsx or .xlsm file.', 'error');
      return;
    }
    setStatus('Reading file…', 'info');
    const reader = new FileReader();
    reader.onload = e => {
      try {
        rawWorkbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const preview = parseDataDictionary(rawWorkbook, '_', '_');
        summaryEl.innerHTML = `
          <strong>${file.name}</strong> loaded successfully.<br>
          <span class="pill">${preview.attributes.length} attributes</span>
          <span class="pill">${preview.refTables.length} reference tables</span>
          <span class="pill">${preview.taxonomy.nodes.length} taxonomy nodes</span>
          <span class="pill">${preview.categoryAttrs.length} category-attribute rows</span>
        `;
        summaryEl.style.display = 'block';
        generateBtn.disabled = false;
        setStatus('', '');
      } catch (err) {
        setStatus('Error reading file: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  async function generate() {
    if (!rawWorkbook) return;
    const tenant = tenantInput.value.trim();
    const taxName = taxNameInput.value.trim();
    const taxDisplay = taxDisplayInput.value.trim() || taxName;

    if (!tenant) { setStatus('Please enter a Syndigo Tenant Name.', 'error'); return; }
    if (!taxName) { setStatus('Please enter a Taxonomy Short Name.', 'error'); return; }

    const deltaMode = deltaToggle.checked;
    if (deltaMode && !modelWorkbook) {
      setStatus('Delta mode needs a current model export — upload one or switch delta mode off.', 'error');
      return;
    }

    generateBtn.disabled = true;
    showWarnings([]);
    showDeltaSummary(null);
    setStatus('Generating model files…', 'info');

    try {
      await new Promise(r => setTimeout(r, 50));
      const data = parseDataDictionary(rawWorkbook, taxName, taxDisplay);

      const zip = new JSZip();

      // Delta mode only emits reference data files — the base models describe structure,
      // not values, so there is nothing to diff in them.
      if (deltaMode) {
        setStatus('Comparing current model against the Data Dictionary…', 'info');
        await tick();
        const { files: deltaFiles, warnings: deltaWarnings, summary } =
          await generateRefDataDelta(data, modelWorkbook);

        showWarnings(deltaWarnings);
        deltaWarnings.forEach(w => console.warn('[ref-delta]', w));

        if (!deltaFiles.length) {
          setStatus('No differences found — the model already matches the Data Dictionary.', 'success');
          generateBtn.disabled = false;
          return;
        }

        for (const { displayName, buf } of deltaFiles) {
          const safeName = displayName.replace(/[\/\\:*?"<>|]/g, '').trim();
          zip.file('080-delta-' + safeName + '.xlsm', buf);
        }

        const totals = summary.reduce((a, s) => ({
          added:      a.added      + s.added,
          updated:    a.updated    + s.updated,
          carried:    a.carried    + s.carried,
          deleted:    a.deleted    + s.deleted,
          relAdded:   a.relAdded   + s.relAdded,
          relDeleted: a.relDeleted + s.relDeleted,
        }), { added: 0, updated: 0, carried: 0, deleted: 0, relAdded: 0, relDeleted: 0 });

        console.table(summary);
        showDeltaSummary(summary, totals);

        setStatus('Packaging zip…', 'info');
        await tick();
        const deltaBlob = await zip.generateAsync({ type: 'blob' });
        saveAs(deltaBlob, `syndigo-refdata-delta-${tenant}-${today()}.zip`);

        let msg = 'Delta ready — ' + deltaFiles.length + ' table(s) changed: ' +
                  totals.added + ' value(s) added, ' + totals.updated + ' updated, ' +
                  totals.deleted + ' deleted, ' + totals.relAdded + ' relationship(s) added, ' +
                  totals.relDeleted + ' removed.';
        if (deltaWarnings.length) msg += ' ' + deltaWarnings.length + ' warning(s) — see browser console.';
        setStatus(msg, deltaWarnings.length ? 'info' : 'success');
        generateBtn.disabled = false;
        return;
      }

      setStatus('Generating 010 — base model (thing)…', 'info');
      await tick();
      const thingBuf = await generateThingModel(data, tenant);
      console.log('[010] type:', Object.prototype.toString.call(thingBuf), 'len:', thingBuf && thingBuf.length);
      zip.file('010-base-model-thing.xlsx', thingBuf);

      setStatus('Generating 040 — base model (reference data)…', 'info');
      await tick();
      const refModelBuf = await generateRefModel(data, tenant);
      console.log('[040] type:', Object.prototype.toString.call(refModelBuf), 'len:', refModelBuf && refModelBuf.length);
      zip.file('040-base-model-reference-data.xlsx', refModelBuf);

      setStatus('Generating 060 — taxonomy model…', 'info');
      await tick();
      const taxBuf = await generateTaxonomyModel(data, tenant);
      console.log('[060] type:', Object.prototype.toString.call(taxBuf), 'len:', taxBuf && taxBuf.length);
      zip.file('060-taxonomy-model.xlsx', taxBuf);

      setStatus('Generating 080 — reference data files (one per table)…', 'info');
      await tick();
      const { files: refFiles, warnings: refWarnings } = await generateRefDataFiles(data);
      for (const { displayName, buf: refBuf } of refFiles) {
        const safeName = displayName.replace(/[\/\\:*?"<>|]/g, '').trim();
        zip.file('080-' + safeName + '.xlsm', refBuf);
      }
      refWarnings.forEach(w => console.warn('[ref-data]', w));
      showWarnings(refWarnings);

      setStatus('Packaging zip…', 'info');
      await tick();
      const blob = await zip.generateAsync({ type: 'blob' });
      saveAs(blob, `syndigo-model-${tenant}-${today()}.zip`);

      var doneMsg = 'Done! Generated ' + refFiles.length + ' reference table file(s).';
      if (refWarnings.length) doneMsg += ' ' + refWarnings.length + ' blank LOV warning(s) — see browser console.';
      setStatus(doneMsg, 'success');
    } catch (err) {
      setStatus('Generation error: ' + err.message, 'error');
      console.error(err);
    } finally {
      generateBtn.disabled = false;
    }
  }
});

function tick() { return new Promise(r => setTimeout(r, 20)); }
function today() { return new Date().toISOString().slice(0, 19).replace('T', '_').replace(/:/g, ''); }

