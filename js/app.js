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
      const thingWb = await generateThingModel(data, tenant);
      zip.file('010-base-model-thing.xlsx', XLSX.write(thingWb, { bookType: 'xlsx', type: 'array' }));

      setStatus('Generating 040 — base model (reference data)…', 'info');
      await tick();
      const refWb = await generateRefModel(data, tenant);
      zip.file('040-base-model-reference-data.xlsx', XLSX.write(refWb, { bookType: 'xlsx', type: 'array' }));

      setStatus('Generating 060 — taxonomy model…', 'info');
      await tick();
      const taxWb = await generateTaxonomyModel(data, tenant);
      zip.file('060-taxonomy-model.xlsx', XLSX.write(taxWb, { bookType: 'xlsx', type: 'array' }));

      setStatus('Generating 080 — reference data values…', 'info');
      await tick();
      const { wb: refDataWb, entityIds } = generateRefData(data);
      zip.file('080-reference-data.xlsm', XLSX.write(refDataWb, { bookType: 'xlsm', type: 'array' }));

      setStatus('Generating 100 — reference data relationships…', 'info');
      await tick();
      const relWb = generateRefRelationships(data, entityIds);
      zip.file('100-reference-data-relationship.xlsm', XLSX.write(relWb, { bookType: 'xlsm', type: 'array' }));

      setStatus('Packaging zip…', 'info');
      await tick();
      const blob = await zip.generateAsync({ type: 'blob' });
      saveAs(blob, `syndigo-model-${tenant}-${today()}.zip`);

      setStatus('Done! Your files have been downloaded.', 'success');
    } catch (err) {
      setStatus('Generation error: ' + err.message, 'error');
      console.error(err);
    } finally {
      generateBtn.disabled = false;
    }
  }
});

function tick() { return new Promise(r => setTimeout(r, 20)); }
function today() { return new Date().toISOString().slice(0, 10); }
