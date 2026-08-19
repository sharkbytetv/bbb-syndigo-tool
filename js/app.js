let rawWorkbook = null;

document.addEventListener('DOMContentLoaded', () => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const tenantInput = document.getElementById('tenantInput');
  const taxNameInput = document.getElementById('taxNameInput');
  const taxDisplayInput = document.getElementById('taxDisplayInput');
  const generateBtn = document.getElementById('generateBtn');
  const statusEl = document.getElementById('status');
  const summaryEl = document.getElementById('summary');

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

  function setStatus(msg, type = 'info') {
    statusEl.textContent = msg;
    statusEl.className = 'status ' + type;
    statusEl.style.display = msg ? 'block' : 'none';
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

    generateBtn.disabled = true;
    setStatus('Generating model files…', 'info');

    try {
      await new Promise(r => setTimeout(r, 50));
      const data = parseDataDictionary(rawWorkbook, taxName, taxDisplay);

      const zip = new JSZip();

      setStatus('Generating 010 — base model (thing)…', 'info');
      await tick();
      const thingBuf = await generateThingModel(data, tenant);
      zip.file('010-base-model-thing.xlsx', thingBuf);

      setStatus('Generating 040 — base model (reference data)…', 'info');
      await tick();
      const refModelBuf = await generateRefModel(data, tenant);
      zip.file('040-base-model-reference-data.xlsx', refModelBuf);

      setStatus('Generating 060 — taxonomy model…', 'info');
      await tick();
      const taxBuf = await generateTaxonomyModel(data, tenant);
      zip.file('060-taxonomy-model.xlsx', taxBuf);

      setStatus('Generating 080 — reference data files (one per table)…', 'info');
      await tick();
      const { files: refFiles, warnings: refWarnings } = await generateRefDataFiles(data);
      for (const { displayName, buf: refBuf } of refFiles) {
        const safeName = displayName.replace(/[\/\\:*?"<>|]/g, '').trim();
        zip.file('080-' + safeName + '.xlsm', refBuf);
      }
      refWarnings.forEach(w => console.warn('[ref-data]', w));

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
