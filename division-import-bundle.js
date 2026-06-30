(function () {

  // ─── Styles ────────────────────────────────────────────────────────────────
  const CSS = `
    .di-wrapper { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 760px; }
    .di-title   { font-size: 20px; font-weight: 600; margin-bottom: 4px; }
    .di-sub     { font-size: 13px; color: #555; margin-bottom: 18px; }
    .di-drop    { border: 2px dashed #aaa; border-radius: 8px; padding: 32px;
                  text-align: center; cursor: pointer; color: #555; margin-bottom: 10px;
                  transition: all .2s; }
    .di-drop.over { border-color: #0078d4; background: #e8f0fe; color: #0078d4; }
    .di-filename { font-size: 12px; color: #444; margin-bottom: 14px; min-height: 16px; }
    .di-options  { display: flex; gap: 10px; align-items: center; margin-bottom: 16px; flex-wrap: wrap; }
    .di-btn     { background: #0078d4; color: #fff; border: none; border-radius: 4px;
                  padding: 9px 22px; font-size: 14px; cursor: pointer; }
    .di-btn:disabled { background: #bbb; cursor: not-allowed; }
    .di-btn.secondary { background: #fff; color: #0078d4; border: 1px solid #0078d4; }
    .di-btn.secondary:disabled { background: #fff; color: #bbb; border-color: #bbb; }
    .di-dry-badge { font-size: 11px; background: #fff4ce; color: #7a4f00;
                    border: 1px solid #c8a400; border-radius: 10px; padding: 2px 8px; }
    .di-log     { margin-top: 18px; max-height: 340px; overflow-y: auto;
                  border: 1px solid #ddd; border-radius: 4px; padding: 10px 12px;
                  font-size: 12.5px; font-family: monospace; line-height: 1.6; }
    .di-row     { border-bottom: 1px solid #f5f5f5; padding: 2px 0; }
    .di-ok      { color: #107c10; }
    .di-skip    { color: #797673; }
    .di-dry     { color: #7a4f00; }
    .di-err     { color: #c50f1f; }
    .di-info    { color: #0078d4; font-weight: 500; }
    .di-head    { color: #323130; font-weight: 600; }
    .di-legend  { font-size: 11px; color: #888; margin-top: 6px; }
    select.di-sel { font-size: 13px; padding: 6px 10px; border: 1px solid #ccc;
                    border-radius: 4px; cursor: pointer; }
  `;

  // ─── Simple CSV Parser ──────────────────────────────────────────────────────
  function parseCSV(text) {
    const lines = text.trim().split(/\r?\n/);
    if (lines.length < 2) return [];
    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    return lines.slice(1)
      .map(line => {
        const vals = line.split(',').map(v => v.trim().replace(/^"|"$/g, ''));
        const obj = {};
        headers.forEach((h, i) => (obj[h] = vals[i] ?? ''));
        return obj;
      })
      .filter(r => Object.values(r).some(v => v !== ''));
  }

  // ─── API Helper ─────────────────────────────────────────────────────────────
  async function apiFetch(path, fetchOptions, token) {
    const res = await fetch(path, {
      ...fetchOptions,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(fetchOptions.headers ?? {}),
      },
    });
    if (res.status === 204) return null;
    const text = await res.text();
    if (!res.ok) throw new Error(`HTTP ${res.status} on ${path}: ${text}`);
    return text ? JSON.parse(text) : null;
  }

  // ─── Find M.PCM.Product by a property value ─────────────────────────────────
  async function findProduct(token, propertyName, value) {
    const body = {
      query: {
        filter: {
          operator: 'and',
          conditions: [
            { operator: 'eq', field: 'definition_name',  value: 'M.PCM.Product' },
            { operator: 'eq', field: propertyName,       value: value }
          ]
        }
      },
      take: 1
    };
    const data = await apiFetch('/api/entities/query', { method: 'POST', body: JSON.stringify(body) }, token);
    return data?.items?.[0] ?? null;
  }

  // ─── Find TB.Division by Identifier OR Label ─────────────────────────────────
  async function findDivision(token, propertyName, value) {
    // propertyName: 'Identifier' | 'Label' | 'Name'
    const body = {
      query: {
        filter: {
          operator: 'and',
          conditions: [
            { operator: 'eq', field: 'definition_name', value: 'TB.Division' },
            { operator: 'eq', field: propertyName,      value: value }
          ]
        }
      },
      take: 1
    };
    const data = await apiFetch('/api/entities/query', { method: 'POST', body: JSON.stringify(body) }, token);
    return data?.items?.[0] ?? null;
  }

  // ─── Get current TB.Division parents of a product ───────────────────────────
  // M.PCM.Product is CHILD → TB.Division is PARENT in TB.PCM.DivisionSelected
  async function getExistingDivisions(token, productId) {
    const data = await apiFetch(
      `/api/entities/${productId}/relations/TB.PCM.DivisionSelected?num_results=500`,
      { method: 'GET' },
      token
    );
    // Parents array contains the TB.Division entities linked to this product
    const parents = data?.parents?.items ?? data?.parents ?? [];
    return parents.map(p => {
      if (typeof p === 'object') {
        return String(p.id ?? p.href?.split('/').pop() ?? '');
      }
      return String(p);
    }).filter(Boolean);
  }

  // ─── Write the merged set back ───────────────────────────────────────────────
  async function setDivisions(token, productId, allDivisionIds) {
    const body = {
      parents: {
        items: allDivisionIds.map(id => ({ href: `/api/entities/${id}` }))
      }
    };
    await apiFetch(
      `/api/entities/${productId}/relations/TB.PCM.DivisionSelected`,
      { method: 'PUT', body: JSON.stringify(body) },
      token
    );
  }

  // ─── Process all CSV rows ────────────────────────────────────────────────────
  async function processRows(rows, token, log, opts) {
    const {
      productField,   // which M.PCM.Product property to search by
      divisionField,  // 'Identifier' | 'Label' | 'Name'
      dryRun          // boolean
    } = opts;

    let added = 0, skipped = 0, errors = 0;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const productVal  = (row['ProductIdentifier'] ?? row[Object.keys(row)[0]] ?? '').trim();
      const divisionVal = (row['DivisionValue']     ?? row[Object.keys(row)[1]] ?? '').trim();

      if (!productVal || !divisionVal) {
        log(`Row ${i + 1}: Skipping — missing value (got: ${JSON.stringify(row)})`, 'di-err');
        errors++;
        continue;
      }

      log(`Row ${i + 1}: Product="${productVal}" → Division "${divisionVal}" …`, 'di-info');

      try {
        // 1. Resolve product
        const product = await findProduct(token, productField, productVal);
        if (!product) {
          log(`  ✗ No M.PCM.Product found where ${productField} = "${productVal}"`, 'di-err');
          errors++;
          continue;
        }

        // 2. Resolve division
        const division = await findDivision(token, divisionField, divisionVal);
        if (!division) {
          log(`  ✗ No TB.Division found where ${divisionField} = "${divisionVal}"`, 'di-err');
          errors++;
          continue;
        }

        const productId  = String(product.id);
        const divisionId = String(division.id);

        // 3. Get existing divisions
        const existing = await getExistingDivisions(token, productId);

        // 4. Check duplicate
        if (existing.includes(divisionId)) {
          log(`  – Division already linked. Skipped.`, 'di-skip');
          skipped++;
          continue;
        }

        // 5. Merge
        const merged = [...existing, divisionId];

        if (dryRun) {
          log(`  [DRY RUN] Would add division (id:${divisionId}). New total: ${merged.length}`, 'di-dry');
        } else {
          await setDivisions(token, productId, merged);
          log(`  ✓ Division added. Total divisions now: ${merged.length}`, 'di-ok');
        }
        added++;

      } catch (err) {
        log(`  ✗ Error: ${err.message}`, 'di-err');
        errors++;
      }
    }

    log('─'.repeat(60), 'di-head');
    if (dryRun) {
      log(`DRY RUN complete — ${added} would be added, ${skipped} already linked, ${errors} errors.`, 'di-dry');
    } else {
      log(`Import complete — ${added} added, ${skipped} already linked, ${errors} errors.`, 'di-head');
    }
  }

  // ─── Render entry point ──────────────────────────────────────────────────────
  async function render(element, options) {
    const config = options?.config ?? {};

    // Defaults — override via Configuration JSON in CH admin
    const defaultProductField  = config.productSearchField  ?? 'Identifier';
    const defaultDivisionField = config.divisionSearchField ?? 'Label';

    // Inject styles
    const style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);

    element.innerHTML = `
      <div class="di-wrapper">
        <div class="di-title">📂 Division Import Tool</div>
        <div class="di-sub">
          Additively links TB.Division records to M.PCM.Product entities.<br/>
          Existing divisions on a product are <strong>never removed</strong> — only new ones are added.
        </div>

        <div class="di-drop" id="di-drop">
          📎 Drop your CSV file here, or <u>click to browse</u>
          <input type="file" id="di-file" accept=".csv" style="display:none"/>
        </div>
        <div class="di-filename" id="di-filename">No file selected</div>

        <div class="di-options">
          <label style="font-size:13px;">
            Match product by:&nbsp;
            <select class="di-sel" id="di-prod-field">
              <option value="Identifier" ${defaultProductField==='Identifier'?'selected':''}>Identifier</option>
              <option value="Label"      ${defaultProductField==='Label'?'selected':''}>Label</option>
              <option value="Name"       ${defaultProductField==='Name'?'selected':''}>Name</option>
            </select>
          </label>
          <label style="font-size:13px;">
            Match division by:&nbsp;
            <select class="di-sel" id="di-div-field">
              <option value="Label"      ${defaultDivisionField==='Label'?'selected':''}>Label</option>
              <option value="Identifier" ${defaultDivisionField==='Identifier'?'selected':''}>Identifier</option>
              <option value="Name"       ${defaultDivisionField==='Name'?'selected':''}>Name</option>
            </select>
          </label>
        </div>

        <div style="display:flex;gap:10px;flex-wrap:wrap;">
          <button class="di-btn secondary" id="di-dry"  disabled>🔍 Dry Run (Preview)</button>
          <button class="di-btn"           id="di-run"  disabled>⬆ Import Divisions</button>
        </div>

        <div class="di-legend">
          CSV columns expected: <code>ProductIdentifier</code>, <code>DivisionValue</code>
        </div>

        <div class="di-log" id="di-log" style="display:none"></div>
      </div>
    `;

    const dropZone  = element.querySelector('#di-drop');
    const fileInput = element.querySelector('#di-file');
    const fileLabel = element.querySelector('#di-filename');
    const dryBtn    = element.querySelector('#di-dry');
    const runBtn    = element.querySelector('#di-run');
    const logPanel  = element.querySelector('#di-log');
    const prodField = element.querySelector('#di-prod-field');
    const divField  = element.querySelector('#di-div-field');

    let parsedRows = [];

    function logMsg(msg, cls = '') {
      logPanel.style.display = 'block';
      const div = document.createElement('div');
      div.className = `di-row ${cls}`;
      div.textContent = msg;
      logPanel.appendChild(div);
      logPanel.scrollTop = logPanel.scrollHeight;
    }

    function loadFile(file) {
      if (!file) return;
      const reader = new FileReader();
      reader.onload = e => {
        parsedRows = parseCSV(e.target.result);
        fileLabel.textContent = `${file.name} — ${parsedRows.length} data row(s) ready`;
        const ok = parsedRows.length > 0;
        dryBtn.disabled = !ok;
        runBtn.disabled = !ok;
      };
      reader.readAsText(file);
    }

    dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('over'); });
    dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('over'));
    dropZone.addEventListener('drop', e => {
      e.preventDefault(); dropZone.classList.remove('over');
      loadFile(e.dataTransfer.files[0]);
    });
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change',   () => loadFile(fileInput.files[0]));

    async function runImport(dryRun) {
      dryBtn.disabled = true;
      runBtn.disabled = true;
      logPanel.innerHTML = '';
      logPanel.style.display = 'block';
      logMsg(dryRun ? '── DRY RUN (no changes will be saved) ──' : '── LIVE IMPORT ──', 'di-head');

      try {
        const token = await options.context.getToken();
        await processRows(parsedRows, token, logMsg, {
          productField:  prodField.value,
          divisionField: divField.value,
          dryRun,
        });
      } catch (err) {
        logMsg(`Fatal: ${err.message}`, 'di-err');
      } finally {
        dryBtn.disabled = false;
        runBtn.disabled = false;
      }
    }

    dryBtn.addEventListener('click', () => runImport(true));
    runBtn.addEventListener('click', () => runImport(false));
  }

  // Export for Content Hub's external component loader
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { default: render };
  } else {
    window.ExternalPageComponent = { default: render };
  }

})();