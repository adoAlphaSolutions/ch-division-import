// ============================================================================
// Division Import Tool — Content Hub External Page Component
// ----------------------------------------------------------------------------
// Reads an Excel (.xlsx) or CSV file IN THE BROWSER and creates one
// TB.PCM.DivisionTempImport staging entity per row. A Content Hub trigger +
// C# action-script then does the additive merge into M.PCM.Product and
// deletes the staging row.
//
// The browser NEVER uploads the file itself — SheetJS parses it in memory and
// only the parsed field values (ProductId, DivisionsCodes) are POSTed.
//
// Expected columns in the spreadsheet (first worksheet):
//   ProductId       -> value written to TB.PCM.DivisionTempImport.ProductId
//   DivisionsCodes  -> value written to TB.PCM.DivisionTempImport.DivisionsCodes
//                      (comma-separated codes, e.g. "AZ,CA,TX")
//
// Configuration textarea (JSON) in the External component:
//   {
//     "definitionName": "TB.PCM.DivisionTempImport",
//     "productField": "id"        // "id" = ProductId cell already holds the CH id.
//                                  // Otherwise a field name (e.g. "Identifier")
//                                  // and the tool resolves it to the CH id first.
//   }
// ============================================================================
 
const XLSX_CDN = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs';
const CH_HOST = window.location.origin;
 
const CSS = `
  .di-wrap   { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 780px; }
  .di-title  { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .di-sub    { font-size: 13px; color: #555; margin-bottom: 18px; }
  .di-drop   { border: 2px dashed #aaa; border-radius: 8px; padding: 32px;
               text-align: center; cursor: pointer; color: #555; margin-bottom: 12px; }
  .di-drop.di-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .di-row    { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; }
  .di-btn    { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer;
               font-size: 14px; }
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
 
let XLSX; // cached SheetJS module
 
async function loadXLSX() {
  if (!XLSX) XLSX = await import(XLSX_CDN);
  return XLSX;
}
 
async function getToken(options) {
  try {
    if (options && options.context && options.context.getToken) {
      return await options.context.getToken();
    }
    if (window.Stylelabs && window.Stylelabs.context && window.Stylelabs.context.getUserToken) {
      return await window.Stylelabs.context.getUserToken();
    }
  } catch (e) { /* fall through */ }
  return null;
}
 
async function apiFetch(path, opts, token) {
  const res = await fetch(`${CH_HOST}${path}`, {
    ...opts,
    headers: {
      'Content-Type': 'application/json',
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(opts && opts.headers ? opts.headers : {})
    }
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status} ${res.statusText}${text ? ' — ' + text : ''}`);
  }
  const ct = res.headers.get('content-type') || '';
  return ct.includes('application/json') ? res.json() : res.text();
}
 
// Read the first worksheet into [{ProductId, DivisionsCodes}, ...], header-tolerant.
async function parseFile(file) {
  const xlsx = await loadXLSX();
  const buf = await file.arrayBuffer();
  const wb = xlsx.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = xlsx.utils.sheet_to_json(ws, { defval: '' });
 
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
 
// If productField is not "id", look up M.PCM.Product by that field and return the CH id.
async function resolveProductId(value, field, token) {
  if (!field || field.toLowerCase() === 'id') return value;
  const q = `Definition.Name=='M.PCM.Product' and ${field}=='${value.replace(/'/g, "\\'")}'`;
  const data = await apiFetch(`/api/entities/query?query=${encodeURIComponent(q)}&take=1`, { method: 'GET' }, token);
  const item = data && data.items && data.items[0];
  return item ? String(item.id) : null;
}
 
async function createStagingRow(definitionName, productId, codes, token) {
  const body = {
    entitydefinition: { href: `${CH_HOST}/api/entitydefinitions/${definitionName}` },
    properties: { ProductId: String(productId), DivisionsCodes: String(codes) }
  };
  const res = await fetch(`${CH_HOST}/api/entities`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: `Bearer ${token}` } : {}) },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`HTTP ${res.status} — ${await res.text().catch(() => '')}`);
  return res.headers.get('Location') || '(created)';
}
 
// ---------------------------------------------------------------------------
export default function render(element, options) {
  const cfg = (options && options.config) || {};
  const definitionName = cfg.definitionName || 'TB.PCM.DivisionTempImport';
  const productField = cfg.productField || 'id';
 
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
      Columns: <b>ProductId</b>, <b>DivisionsCodes</b> (comma-separated, e.g. AZ,CA,TX)
    </div>
  `;
 
  element.innerHTML = '';
  element.appendChild(style);
  element.appendChild(wrap);
 
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
 
    const token = await getToken(options);
    if (!token && !dryRun) log('⚠ No token from context — POSTs may be rejected.', 'di-err');
 
    let created = 0, errors = 0;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const label = `Row ${i + 1}: ProductId="${r.ProductId}" Divisions="${r.DivisionsCodes}"`;
      try {
        if (!r.ProductId || !r.DivisionsCodes) {
          log(`${label} — missing ProductId or DivisionsCodes. Skipped.`, 'di-err');
          errors++; continue;
        }
        let pid = r.ProductId;
        if (productField.toLowerCase() !== 'id') {
          pid = await resolveProductId(r.ProductId, productField, token);
          if (!pid) { log(`${label} — no M.PCM.Product where ${productField}="${r.ProductId}".`, 'di-err'); errors++; continue; }
        }
        if (dryRun) {
          log(`${label} → would create ${definitionName} (ProductId=${pid})`, 'di-skip');
        } else {
          const loc = await createStagingRow(definitionName, pid, r.DivisionsCodes, token);
          log(`${label} → created ✓ ${loc}`, 'di-ok');
          created++;
        }
      } catch (e) {
        log(`${label} → error: ${e.message}`, 'di-err');
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
}
