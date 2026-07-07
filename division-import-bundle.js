// ============================================================================
// Division Import Tool — Content Hub External Component (CH 4.2+ / React UI)
// ----------------------------------------------------------------------------
// Reads an Excel (.xlsx) or CSV file IN THE BROWSER and creates one
// TB.PCM.DivisionTempImport staging entity per row, using the pre-authenticated
// Content Hub JavaScript SDK client that Content Hub passes into render(context).
// A Content Hub trigger + C# action-script then merges the divisions additively
// into M.PCM.Product and deletes the staging row.
//
// IMPORTANT — bundle contract (CH 4.2+):
//   The default export is called as createExternalRoot(rootElement) and must
//   return an object with render(context) and unmount(). The old
//   `export default function render(element, options)` shape does NOT receive
//   the authenticated client — which is why the token was missing before.
//
// context provides:
//   context.client   -> pre-authenticated CH JavaScript SDK client
//   context.config   -> JSON object from the Configuration textarea
//   context.culture  -> current culture (e.g. "en-US")
//
// Spreadsheet columns (first worksheet):
//   ProductId       -> the internal Content Hub id of the M.PCM.Product (e.g. 2880931)
//   DivisionsCodes  -> comma-separated division codes (e.g. "AZ,CA,TX")
//
// Configuration textarea (JSON), optional:
//   { "definitionName": "TB.PCM.DivisionTempImport" }
// ============================================================================

const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

const CSS = `
  .di-wrap   { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 780px; }
  .di-title  { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .di-sub    { font-size: 13px; color: #555; margin-bottom: 18px; }
  .di-drop   { border: 2px dashed #aaa; border-radius: 8px; padding: 32px;
               text-align: center; cursor: pointer; color: #555; margin-bottom: 12px; }
  .di-drop.di-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .di-row    { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; }
  .di-btn    { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .di-btn:disabled { opacity: .5; cursor: not-allowed; }
  .di-dry    { background: #edf2f7; color: #2d3748; }
  .di-go     { background: #2b6cb0; color: #fff; }
  .di-log    { background: #1a202c; color: #e2e8f0; font-family: monospace; font-size: 12px;
               padding: 14px; border-radius: 6px; margin-top: 14px; max-height: 320px;
               overflow: auto; white-space: pre-wrap; display: none; }
  .di-ok   { color: #68d391; }
  .di-skip { color: #cbd5e0; }
  .di-err  { color: #fc8181; }
  .di-info { color: #90cdf4; }
`;

// Load SheetJS (UMD build → window.XLSX). Script tag is more CSP-friendly than a
// cross-origin dynamic import().
function loadXLSX() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) return resolve(window.XLSX);
    const s = document.createElement('script');
    s.src = SHEETJS_URL;
    s.onload = () => (window.XLSX ? resolve(window.XLSX) : reject(new Error('SheetJS loaded but window.XLSX is missing.')));
    s.onerror = () => reject(new Error('Could not load SheetJS (check CSP / network access).'));
    document.head.appendChild(s);
  });
}

// Read first worksheet into [{ProductId, DivisionsCodes}, ...], header-tolerant.
async function parseFile(file) {
  const XLSX = await loadXLSX();
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });

  return raw.map(r => {
    const out = {};
    for (const key of Object.keys(r)) {
      const k = key.trim().toLowerCase();
      if (k === 'productid') out.ProductId = String(r[key]).trim();
      else if (k === 'divisionscodes' || k === 'divisioncodes') out.DivisionsCodes = String(r[key]).trim();
    }
    return out;
  }).filter(r => r.ProductId || r.DivisionsCodes);
}

// Set a property whether or not it is localized: try plain, fall back to culture.
function setProp(entity, name, value, culture) {
  try {
    entity.setPropertyValue(name, value);
  } catch (e) {
    entity.setPropertyValue(name, value, culture);
  }
}

// Create one staging entity via the authenticated SDK client.
async function createStagingRow(client, definitionName, productId, codes, culture) {
  const entity = await client.entityFactory.createAsync(definitionName);
  setProp(entity, 'ProductId', String(productId), culture);
  setProp(entity, 'DivisionsCodes', String(codes), culture);
  const saved = await client.entities.saveAsync(entity);
  return (saved && (saved.id || saved.Id)) || entity.id || '(created)';
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render(context) {
      const client = context && context.client;
      const cfg = (context && context.config) || {};
      const culture = (context && context.culture) || 'en-US';
      const definitionName = cfg.definitionName || 'TB.PCM.DivisionTempImport';

      const style = document.createElement('style');
      style.textContent = CSS;

      const wrap = document.createElement('div');
      wrap.className = 'di-wrap';
      wrap.innerHTML = `
        <div class="di-title">📂 Division Import Tool</div>
        <div class="di-sub">Uploads a spreadsheet and stages each row as a
          ${definitionName} entity. Divisions are merged additively by the action-script;
          existing divisions are never removed.</div>
        <div class="di-drop" id="di-drop">📎 Drop your .xlsx / .csv file here, or click to browse</div>
        <input type="file" id="di-file" accept=".xlsx,.xls,.csv" style="display:none" />
        <div class="di-row">
          <button class="di-btn di-dry" id="di-dry" disabled>🔍 Dry Run</button>
          <button class="di-btn di-go"  id="di-go"  disabled>⬆ Create staging rows</button>
          <span id="di-status" style="font-size:13px;color:#555"></span>
        </div>
        <div class="di-log" id="di-log"></div>
        <div style="font-size:12px;color:#888;margin-top:10px">
          Columns: <b>ProductId</b> (Content Hub id), <b>DivisionsCodes</b> (comma-separated, e.g. AZ,CA,TX)
        </div>
      `;

      rootElement.innerHTML = '';
      rootElement.appendChild(style);
      rootElement.appendChild(wrap);

      const drop   = wrap.querySelector('#di-drop');
      const input  = wrap.querySelector('#di-file');
      const dryBtn = wrap.querySelector('#di-dry');
      const goBtn  = wrap.querySelector('#di-go');
      const status = wrap.querySelector('#di-status');
      const logEl  = wrap.querySelector('#di-log');

      let rows = [];

      function log(msg, cls) {
        logEl.style.display = 'block';
        const line = document.createElement('div');
        if (cls) line.className = cls;
        line.textContent = msg;
        logEl.appendChild(line);
        logEl.scrollTop = logEl.scrollHeight;
      }
      function clearLog() { logEl.innerHTML = ''; logEl.style.display = 'none'; }

      async function onFile(file) {
        if (!file) return;
        clearLog();
        status.textContent = 'Parsing…';
        try {
          rows = await parseFile(file);
          status.textContent = `${file.name} — ${rows.length} row(s) ready`;
          dryBtn.disabled = rows.length === 0;
          goBtn.disabled = rows.length === 0;
        } catch (e) {
          status.textContent = '';
          log(`✗ Could not parse file: ${e.message}`, 'di-err');
        }
      }

      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('di-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('di-hover'));
      drop.addEventListener('drop', e => {
        e.preventDefault(); drop.classList.remove('di-hover');
        onFile(e.dataTransfer.files[0]);
      });

      async function run(dryRun) {
        clearLog();
        dryBtn.disabled = true; goBtn.disabled = true;
        log(dryRun ? '── DRY RUN (nothing is written) ──' : '── CREATING STAGING ROWS ──', 'di-info');

        if (!dryRun && !client) {
          log('✗ No SDK client in context — the component must run inside Content Hub 4.2+ as an External Component.', 'di-err');
          dryBtn.disabled = false; goBtn.disabled = false;
          return;
        }

        let created = 0, errors = 0;
        for (let i = 0; i < rows.length; i++) {
          const r = rows[i];
          const label = `Row ${i + 1}: ProductId="${r.ProductId}" Divisions="${r.DivisionsCodes}"`;
          try {
            if (!r.ProductId || !r.DivisionsCodes) {
              log(`${label} — missing ProductId or DivisionsCodes. Skipped.`, 'di-err');
              errors++; continue;
            }
            if (dryRun) {
              log(`${label} → would create ${definitionName}`, 'di-skip');
            } else {
              const id = await createStagingRow(client, definitionName, r.ProductId, r.DivisionsCodes, culture);
              log(`${label} → created ✓ (id ${id})`, 'di-ok');
              created++;
            }
          } catch (e) {
            log(`${label} → error: ${e && e.message ? e.message : e}`, 'di-err');
            errors++;
          }
        }

        log('──────────────────────────────────────────', 'di-info');
        log(dryRun
          ? `DRY RUN done — ${rows.length} row(s) would be staged, ${errors} error(s).`
          : `Done — ${created} staging row(s) created, ${errors} error(s). The action-script will now merge divisions.`,
          errors ? 'di-err' : 'di-ok');

        dryBtn.disabled = false; goBtn.disabled = false;
      }

      dryBtn.addEventListener('click', () => run(true));
      goBtn.addEventListener('click', () => run(false));
    },

    unmount() {
      rootElement.innerHTML = '';
    }
  };
}
