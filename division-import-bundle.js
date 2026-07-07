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
//   DivisionsCodes  -> comma-separated division codes (e.g. "AZ,CA,TX").
//                      Non-empty = ADD those divisions. Empty = REMOVE the pattern.
//
// The bundle derives the PATTERN = the union of all non-empty DivisionsCodes cells
// in the file, and stamps it onto the PatternCodes field of EVERY staging row, so a
// per-row action-script can tell an empty row what to remove.
//
// Requires these fields on TB.PCM.DivisionTempImport:
//   ProductId (String), DivisionsCodes (String), PatternCodes (String)
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

// Create a local entity. Some Content Hub SDK versions accept only the
// definition name; others require a culture / load option as a 2nd argument
// (and internally call .toString() on it — passing nothing yields the
// "reading 'toString' of undefined" error). Try the variants in order.
async function createEntity(client, definitionName, culture) {
  const attempts = [
    () => client.entityFactory.createAsync(definitionName),
    () => client.entityFactory.createAsync(definitionName, culture),
    () => client.entityFactory.createAsync(definitionName, [culture])
  ];
  let lastErr;
  for (const attempt of attempts) {
    try {
      const entity = await attempt();
      if (entity) return entity;
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error(`entityFactory.createAsync failed: ${lastErr && lastErr.message ? lastErr.message : lastErr}`);
}

// Best-effort enumeration of the property names actually present on an entity,
// so we can see whether e.g. "PatternCodes" exists and its exact casing.
function listPropNames(entity) {
  try {
    const p = entity && entity.properties;
    if (Array.isArray(p)) return p.map(x => x.name || x.Name).filter(Boolean).join(', ');
    if (p && typeof p === 'object') return Object.keys(p).join(', ');
    if (entity && typeof entity.getProperties === 'function') {
      const arr = entity.getProperties();
      if (Array.isArray(arr)) return arr.map(x => x.name || x.Name).filter(Boolean).join(', ');
    }
  } catch (e) {
    return `(could not list: ${e && e.message ? e.message : e})`;
  }
  return '(unknown)';
}

// Set a property whether or not it is localized: try plain, then with culture.
function setProp(entity, name, value, culture) {
  try {
    entity.setPropertyValue(name, value);
  } catch (e1) {
    try {
      entity.setPropertyValue(name, value, culture);
    } catch (e2) {
      throw new Error(`setPropertyValue('${name}') failed: ${e1 && e1.message ? e1.message : e1} | ${e2 && e2.message ? e2.message : e2}`);
    }
  }
}

// Create one staging entity via the authenticated SDK client.
async function createStagingRow(client, definitionName, productId, codes, patternCodes, culture) {
  const entity = await createEntity(client, definitionName, culture);
  try {
    setProp(entity, 'ProductId', String(productId), culture);
    setProp(entity, 'DivisionsCodes', String(codes || ''), culture);
    setProp(entity, 'PatternCodes', String(patternCodes || ''), culture);
  } catch (e) {
    throw new Error(`${e && e.message ? e.message : e}. Entity properties present: [${listPropNames(entity)}]`);
  }
  let saved;
  try {
    saved = await client.entities.saveAsync(entity);
  } catch (e) {
    throw new Error(`saveAsync failed: ${e && e.message ? e.message : e}`);
  }
  return (saved && (saved.id || saved.Id)) || entity.id || '(created)';
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render(context) {
      const client = context && context.client;
      const cfg = (context && context.config) || {};
      const culture = (context && (context.culture || (context.options && context.options.culture))) || 'en-US';
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
          Columns: <b>ProductId</b> (CH id), <b>DivisionsCodes</b> (comma-separated = add; empty = remove the pattern).
          The pattern is auto-detected and written to <b>PatternCodes</b> on every row.
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

        // Pattern = union of all non-empty DivisionsCodes cells in the file.
        const patternSet = new Set();
        for (const r of rows) {
          if (r.DivisionsCodes) {
            r.DivisionsCodes.split(',').map(c => c.trim()).filter(Boolean).forEach(c => patternSet.add(c));
          }
        }
        const patternCodes = Array.from(patternSet).join(',');
        log(`Pattern detected: ${patternCodes || '(none)'}`, 'di-info');
        if (!patternCodes) {
          log('⚠ No non-empty DivisionsCodes in the file — empty rows will have nothing to remove.', 'di-err');
        }

        let created = 0, errors = 0;
        for (let i = 0; i < rows.length; i++) {
          const r = rows[i];
          const op = r.DivisionsCodes ? 'ADD' : 'REMOVE';
          const shown = op === 'ADD' ? r.DivisionsCodes : patternCodes;
          const label = `Row ${i + 1}: ProductId="${r.ProductId}" [${op} ${shown}]`;
          try {
            if (!r.ProductId) {
              log(`${label} — missing ProductId. Skipped.`, 'di-err');
              errors++; continue;
            }
            if (dryRun) {
              log(`${label} → would create ${definitionName} (PatternCodes="${patternCodes}")`, 'di-skip');
            } else {
              const id = await createStagingRow(client, definitionName, r.ProductId, r.DivisionsCodes, patternCodes, culture);
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
