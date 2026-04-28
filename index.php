<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Prodboard Project Reader</title>
  <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: system-ui, -apple-system, sans-serif;
      font-size: 14px;
      background: #f5f5f5;
      color: #1a1a1a;
    }

    header {
      background: #1a1a2e;
      color: #fff;
      padding: 16px 24px;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    header h1 { font-size: 18px; font-weight: 600; }
    header .subtitle { font-size: 12px; color: #aaa; margin-top: 2px; }

    .container { max-width: 1400px; margin: 0 auto; padding: 24px; }

    .tabs-nav {
      display: flex;
      gap: 0;
      border-bottom: 2px solid #e5e7eb;
      margin-bottom: 20px;
    }
    .tab-btn {
      padding: 10px 22px;
      background: none;
      border: none;
      border-bottom: 2px solid transparent;
      margin-bottom: -2px;
      font-size: 13px;
      font-weight: 500;
      color: #6b7280;
      cursor: pointer;
      transition: color .15s, border-color .15s;
    }
    .tab-btn:hover { color: #1a1a1a; }
    .tab-btn.active { color: #4f46e5; border-bottom-color: #4f46e5; font-weight: 600; }
    .tab-pane { display: none; }
    .tab-pane.active { display: block; }

    .card {
      background: #fff;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 20px;
    }
    .card h2 { font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: .05em; color: #666; margin-bottom: 16px; }

    .form-grid {
      display: grid;
      gap: 12px;
    }
    .form-grid.cols-2 { grid-template-columns: repeat(2, 1fr); }
    .form-grid.cols-3 { grid-template-columns: repeat(3, 1fr); }
    .form-grid.cols-4 { grid-template-columns: repeat(4, 1fr); }

    label { display: block; font-size: 12px; font-weight: 500; color: #555; margin-bottom: 4px; }

    input[type=text], input[type=password], input[type=date], select {
      width: 100%;
      padding: 8px 10px;
      border: 1px solid #d0d0d0;
      border-radius: 6px;
      font-size: 13px;
      background: #fafafa;
      transition: border-color .15s;
      outline: none;
    }
    input:focus, select:focus { border-color: #4f46e5; background: #fff; }

    .btn {
      padding: 9px 20px;
      border: none;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
      transition: opacity .15s;
    }
    .btn:disabled { opacity: .5; cursor: not-allowed; }
    .btn-primary { background: #4f46e5; color: #fff; }
    .btn-primary:hover:not(:disabled) { background: #4338ca; }
    .btn-secondary { background: #e5e7eb; color: #1a1a1a; }
    .btn-secondary:hover:not(:disabled) { background: #d1d5db; }

    .actions { display: flex; gap: 8px; align-items: center; margin-top: 16px; }

    .status-badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 600;
    }
    .status-0 { background: #fef3c7; color: #92400e; }
    .status-1 { background: #dbeafe; color: #1e40af; }
    .status-2 { background: #d1fae5; color: #065f46; }
    .status-3 { background: #f3f4f6; color: #4b5563; }

    .alert {
      padding: 10px 14px;
      border-radius: 6px;
      font-size: 13px;
      margin-bottom: 16px;
    }
    .alert-error { background: #fee2e2; color: #991b1b; border: 1px solid #fecaca; }
    .alert-success { background: #d1fae5; color: #065f46; border: 1px solid #a7f3d0; }
    .alert-info { background: #eff6ff; color: #1e40af; border: 1px solid #bfdbfe; }

    .results-meta {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }
    .results-meta .count { font-size: 13px; color: #666; }

    .table-wrap { overflow-x: auto; }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    thead th {
      background: #f9fafb;
      border-bottom: 2px solid #e5e7eb;
      padding: 10px 12px;
      text-align: left;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .05em;
      color: #6b7280;
      white-space: nowrap;
    }
    tbody tr { border-bottom: 1px solid #f0f0f0; }
    tbody tr:hover { background: #fafafa; }
    tbody td { padding: 10px 12px; vertical-align: middle; }
    tbody td.muted { color: #9ca3af; }

    .spinner {
      width: 20px; height: 20px;
      border: 2px solid #e5e7eb;
      border-top-color: #4f46e5;
      border-radius: 50%;
      animation: spin .7s linear infinite;
      display: inline-block;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    .token-status {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
    }
    .dot { width: 8px; height: 8px; border-radius: 50%; }
    .dot-green { background: #22c55e; }
    .dot-red { background: #ef4444; }
    .dot-gray { background: #d1d5db; }

    .progress-bar {
      height: 4px;
      background: #e5e7eb;
      border-radius: 2px;
      overflow: hidden;
      margin-bottom: 8px;
    }
    .progress-fill {
      height: 100%;
      background: #4f46e5;
      transition: width .3s;
    }

    tbody tr.clickable { cursor: pointer; }
    tbody tr.clickable:hover { background: #f0f4ff; }
    tbody tr.active { background: #eff0ff !important; }

    /* Side panel */
    .panel-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,.3);
      z-index: 100;
    }
    .panel-overlay.open { display: block; }

    .side-panel {
      position: fixed;
      top: 0; right: 0;
      width: min(600px, 100vw);
      height: 100vh;
      background: #fff;
      box-shadow: -4px 0 24px rgba(0,0,0,.15);
      z-index: 101;
      display: flex;
      flex-direction: column;
      transform: translateX(100%);
      transition: transform .25s ease;
    }
    .side-panel.open { transform: translateX(0); }

    .panel-header {
      display: flex;
      align-items: flex-start;
      gap: 12px;
      padding: 20px 24px;
      border-bottom: 1px solid #e5e7eb;
      flex-shrink: 0;
    }
    .panel-header .panel-title { flex: 1; }
    .panel-header h3 { font-size: 16px; font-weight: 600; margin-bottom: 4px; }
    .panel-header .panel-sub { font-size: 12px; color: #9ca3af; }

    .panel-close {
      background: none;
      border: none;
      font-size: 20px;
      cursor: pointer;
      color: #9ca3af;
      line-height: 1;
      padding: 2px 4px;
      flex-shrink: 0;
    }
    .panel-close:hover { color: #1a1a1a; }

    .panel-body {
      overflow-y: auto;
      flex: 1;
      padding: 24px;
    }

    .detail-section { margin-bottom: 24px; }
    .detail-section h4 {
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .05em;
      color: #9ca3af;
      margin-bottom: 10px;
      padding-bottom: 6px;
      border-bottom: 1px solid #f0f0f0;
    }

    .detail-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px 16px;
    }
    .detail-item { }
    .detail-item .di-label { font-size: 11px; color: #9ca3af; margin-bottom: 2px; }
    .detail-item .di-value { font-size: 13px; font-weight: 500; word-break: break-word; }

    .invoice-block {
      border: 1px solid #e5e7eb;
      border-radius: 6px;
      margin-bottom: 12px;
      overflow: hidden;
    }
    .invoice-table-wrap { overflow-x: auto; }
    .invoice-block-header {
      background: #f9fafb;
      padding: 8px 12px;
      font-size: 12px;
      font-weight: 600;
      display: flex;
      justify-content: space-between;
    }
    .invoice-lines { width: 100%; border-collapse: collapse; font-size: 12px; }
    .invoice-lines th {
      background: #f9fafb;
      padding: 6px 10px;
      text-align: left;
      font-size: 10px;
      font-weight: 600;
      text-transform: uppercase;
      color: #9ca3af;
      border-bottom: 1px solid #e5e7eb;
    }
    .invoice-lines td { padding: 6px 10px; border-bottom: 1px solid #f0f0f0; }
    .invoice-lines tr:last-child td { border-bottom: none; }
    .invoice-totals {
      border-top: 1px solid #e5e7eb;
      padding: 8px 12px;
      font-size: 12px;
      display: flex;
      flex-direction: column;
      align-items: flex-end;
      gap: 3px;
    }
    .invoice-totals .total-row { display: flex; gap: 24px; }
    .invoice-totals .total-row strong { font-weight: 600; }

    .tag-pill {
      display: inline-block;
      background: #f3f4f6;
      border-radius: 4px;
      padding: 2px 8px;
      font-size: 12px;
      margin: 2px;
    }

    .panel-loading {
      display: flex;
      align-items: center;
      justify-content: center;
      height: 200px;
      gap: 10px;
      color: #9ca3af;
      font-size: 13px;
    }

    @media (max-width: 768px) {
      .form-grid.cols-2,
      .form-grid.cols-3,
      .form-grid.cols-4 { grid-template-columns: 1fr; }
      .detail-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>

<header>
<img src="logoc.png" alt="Prodboard" style="height:36px;display:block;">  
<img src="logo_png24_big.png" alt="Prodboard" style="height:36px;display:block;">
  
  <div>
    <h1>Prodboard Project Reader</h1>
    <div class="subtitle">__________________________________________</div>
  </div>
  <div style="margin-left:auto" id="tokenStatus">
    <div class="token-status">
      <div class="dot dot-gray" id="tokenDot"></div>
      <span id="tokenLabel">Non connecté</span>
    </div>
  </div>
</header>

<div class="container">

  <!-- Auth -->
  <div class="card">
    <h2>Authentification</h2>
    <div style="display:flex;align-items:flex-end;gap:12px;flex-wrap:wrap">
      <div style="flex:0 0 auto;width:calc(25% - 8px);min-width:100px">
        <label for="company">Entreprise</label>
        <input type="text" readonly id="company" placeholder="votre-slug-entreprise" value="cuisineandcie" autocomplete="off">
      </div>
      <div style="flex:0 0 auto;width:calc(33% - 8px);min-width:140px">
        <label for="privateKey">Clé privée</label>
        <input type="text" id="privateKey" placeholder="Token" value="Wcrl0XYx1obQkcBWMXUETOzCW15uGgej" autocomplete="off">
      </div>
      <div style="display:flex;gap:8px;align-items:center;flex-shrink:0">
        <button class="btn btn-primary" id="btnConnect" onclick="connect()">Se connecter</button>
        <button class="btn btn-secondary" id="btnDisconnect" onclick="disconnect()" style="display:none">Se déconnecter</button>
        <div id="authMsg"></div>
      </div>
    </div>
  </div>

  <!-- Tabs -->
  <div class="tabs-nav">
    <button class="tab-btn active" onclick="switchTab('projets')">Projets</button>
    <button class="tab-btn" onclick="switchTab('config')">Configuration</button>
  </div>

  <!-- Tab: Projets -->
  <div class="tab-pane active" id="tab-projets">

  <!-- Filters -->
  <div class="card">
    <h2>Filtres</h2>
    <div class="form-grid cols-4">
      <div>
        <label for="fStatus">Statut</label>
        <select id="fStatus">
          <option value="">Tous les statuts</option>
          <option value="0">Planification</option>
          <option value="1">Commandé</option>
          <option value="2">Confirmé</option>
          <option value="3">Terminé</option>
        </select>
      </div>
      <div>
        <label for="fCustomer">Client</label>
        <input type="text" id="fCustomer" placeholder="Identité client">
      </div>
      <div>
        <label for="fSubdivision">Subdivision</label>
        <input type="text" id="fSubdivision" placeholder="Code subdivision">
      </div>
      <div>
        <label for="fInstance">Instance</label>
        <input type="text" id="fInstance" placeholder="Code instance">
      </div>
      <div>
        <label for="fReference">Référence</label>
        <input type="text" id="fReference" placeholder="Référence projet">
      </div>
      <div>
        <label for="fTag">Étiquette</label>
        <input type="text" id="fTag" placeholder="Code étiquette">
      </div>
      <div>
        <label for="fDateFrom">Créé à partir du</label>
        <input type="date" id="fDateFrom">
      </div>
      <div>
        <label for="fDateTo">Créé jusqu'au</label>
        <input type="date" id="fDateTo">
      </div>
    </div>
    <div class="actions">
      <button class="btn btn-primary" id="btnSearch" onclick="search()" disabled>Rechercher</button>
      <button class="btn btn-secondary" onclick="clearFilters()">Effacer les filtres</button>
    </div>
  </div>

  <!-- Results -->
  <div class="card" id="resultsCard" style="display:none">
    <div id="progressWrap" style="display:none">
      <div class="progress-bar"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
      <div id="progressLabel" style="font-size:12px;color:#666;margin-bottom:12px;"></div>
    </div>

    <div id="alertBox"></div>

    <div class="results-meta">
      <h2 style="margin:0">Projets</h2>
      <div style="display:flex;align-items:center;gap:12px">
        <span class="count" id="resultCount"></span>
        <button class="btn btn-secondary" id="btnExcel" onclick="downloadExcel()" style="display:none">⬇ Export Excel</button>
      </div>
    </div>

    <div class="table-wrap">
      <table id="resultsTable">
        <thead>
          <tr>
            <th>#</th>
            <th>Référence</th>
            <th>Nom</th>
            <th>Statut</th>
            <th>Client</th>
            <th>Subdivision</th>
            <th>Étiquettes</th>
            <th>Créé le</th>
            <th>Modifié le</th>
            <th style="text-align:right">Total facturé</th>
          </tr>
        </thead>
        <tbody id="resultsBody"></tbody>
      </table>
    </div>
  </div>

  </div><!-- /tab-projets -->

  <!-- Tab: Configuration -->
  <div class="tab-pane" id="tab-config">
    <div class="card">
      <h2>Configuration</h2>
      <p style="font-size:13px;color:#9ca3af;">Aucun paramètre de configuration pour l'instant.</p>
    </div>
  </div><!-- /tab-config -->

</div>

<!-- Detail panel -->
<div class="panel-overlay" id="panelOverlay" onclick="closePanel()"></div>
<div class="side-panel" id="sidePanel">
  <div class="panel-header">
    <div class="panel-title">
      <h3 id="panelTitle">—</h3>
      <div class="panel-sub" id="panelSub"></div>
    </div>
    <button class="panel-close" onclick="closePanel()">✕</button>
  </div>
  <div class="panel-body" id="panelBody">
    <div class="panel-loading"><div class="spinner"></div> Chargement…</div>
  </div>
</div>

<script>
  const BASE = 'https://api-v2.prodboard.com';
  let token = null;
  let currentItems = [];

  // ── Auth ──────────────────────────────────────────────────────────────────

  async function connect() {
    const company    = document.getElementById('company').value.trim();
    const privateKey = document.getElementById('privateKey').value.trim();
    if (!company || !privateKey) { showAuthMsg('error', 'Veuillez saisir l\'entreprise et la clé privée.'); return; }

    const btn = document.getElementById('btnConnect');
    btn.disabled = true;
    btn.textContent = 'Connexion…';
    showAuthMsg('', '');

    try {
      const res = await fetch(`${BASE}/security/get-token-value`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ company, privateKey })
      });

      if (!res.ok) throw new Error(`HTTP ${res.status}`);

      const raw = await res.text();
      // API returns a quoted JSON string, e.g. "\"eyJ...\"" or just "eyJ..."
      token = raw.replace(/^"|"$/g, '').trim();
      if (!token) throw new Error('Jeton vide reçu.');

      setConnected(true);
      showAuthMsg('success', 'Connexion réussie.');
    } catch (e) {
      token = null;
      setConnected(false);
      showAuthMsg('error', `Échec de connexion : ${e.message}`);
    } finally {
      btn.disabled = false;
      btn.textContent = 'Se connecter';
    }
  }

  function disconnect() {
    token = null;
    setConnected(false);
    document.getElementById('authMsg').innerHTML = '';
    document.getElementById('resultsCard').style.display = 'none';
  }

  function setConnected(ok) {
    document.getElementById('tokenDot').className = `dot ${ok ? 'dot-green' : 'dot-gray'}`;
    document.getElementById('tokenLabel').textContent = ok ? 'Connecté' : 'Non connecté';
    document.getElementById('btnDisconnect').style.display = ok ? '' : 'none';
    document.getElementById('btnConnect').style.display = ok ? 'none' : '';
    document.getElementById('btnSearch').disabled = !ok;
  }

  function showAuthMsg(type, text) {
    const el = document.getElementById('authMsg');
    if (!text) { el.innerHTML = ''; return; }
    const cls = type === 'error' ? 'alert-error' : 'alert-success';
    el.innerHTML = `<div class="alert ${cls}" style="margin:0">${text}</div>`;
  }

  // ── Search ────────────────────────────────────────────────────────────────

  async function search() {
    if (!token) return;

    const params = {};
    const status     = document.getElementById('fStatus').value;
    const customer   = document.getElementById('fCustomer').value.trim();
    const subdivision= document.getElementById('fSubdivision').value.trim();
    const instance   = document.getElementById('fInstance').value.trim();
    const reference  = document.getElementById('fReference').value.trim();
    const tag        = document.getElementById('fTag').value.trim();
    const dateFrom   = document.getElementById('fDateFrom').value;
    const dateTo     = document.getElementById('fDateTo').value;

    if (status !== '')  params.status      = status;
    if (customer)       params.customer    = customer;
    if (subdivision)    params.subdivision = subdivision;
    if (instance)       params.instance   = instance;
    if (reference)      params.reference  = reference;
    if (tag)            params.tag        = tag;

    const btn = document.getElementById('btnSearch');
    btn.disabled = true;

    const card = document.getElementById('resultsCard');
    card.style.display = '';
    document.getElementById('alertBox').innerHTML = '';
    document.getElementById('resultsBody').innerHTML = '';
    document.getElementById('resultCount').textContent = '';
    showProgress(true, 0, 'Chargement…');

    try {
      const allItems = await fetchAllPages(params, (fetched, total) => {
        const pct = total > 0 ? Math.round((fetched / total) * 100) : 0;
        showProgress(true, pct, `Récupéré ${fetched} / ${total} projets…`);
      });

      showProgress(false);

      // Client-side date filter
      const from = dateFrom ? new Date(dateFrom) : null;
      const to   = dateTo   ? new Date(dateTo + 'T23:59:59') : null;
      const items = allItems.filter(p => {
        const created = new Date(p.createdAt);
        if (from && created < from) return false;
        if (to   && created > to)   return false;
        return true;
      });

      currentItems = items;
      renderTable(items);
      document.getElementById('resultCount').textContent = `${items.length} projet${items.length !== 1 ? 's' : ''}${allItems.length !== items.length ? ` (filtré sur ${allItems.length})` : ''}`;
      document.getElementById('btnExcel').style.display = items.length > 0 ? '' : 'none';

      if (items.length === 0) {
        document.getElementById('alertBox').innerHTML = '<div class="alert alert-info">Aucun projet trouvé correspondant aux critères.</div>';
      }
    } catch (e) {
      showProgress(false);
      document.getElementById('alertBox').innerHTML = `<div class="alert alert-error">Erreur : ${e.message}</div>`;
    } finally {
      btn.disabled = false;
    }
  }

  async function fetchAllPages(params, onProgress) {
    const headers = { 'Authorization': `Bearer ${token}` };
    let pageIndex = 0;
    let total = null;
    const all = [];

    while (true) {
      const qs = new URLSearchParams({ ...params, pageIndex });
      const res = await fetch(`${BASE}/projects?${qs}`, { headers });
      if (!res.ok) {
        const body = await res.text().catch(() => '');
        throw new Error(`HTTP ${res.status}${body ? ': ' + body.slice(0, 200) : ''}`);
      }
      const data = await res.json();
      total = data.total ?? total;
      all.push(...(data.items ?? []));
      onProgress(all.length, total ?? all.length);

      if (!data.items || data.items.length < 100) break;
      pageIndex++;
    }

    return all;
  }

  // ── Render ────────────────────────────────────────────────────────────────

  const STATUS_LABEL = ['Planification', 'Commandé', 'Confirmé', 'Terminé'];

  function renderTable(items) {
    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';

    items.forEach(p => {
      const invoiceTotal = computeInvoiceTotal(p);
      const tr = document.createElement('tr');
      tr.className = 'clickable';
      tr.dataset.guid = p.guid;
      tr.addEventListener('click', () => openPanel(p));
      tr.innerHTML = `
        <td>${esc(p.number ?? '')}</td>
        <td>${esc(p.reference ?? '')}</td>
        <td><strong>${esc(p.name ?? '')}</strong></td>
        <td>${statusBadge(p.status)}</td>
        <td>${esc(p.customer?.name ?? p.customer?.identity ?? '')}</td>
        <td class="muted">${esc(p.subdivision ?? '')}</td>
        <td class="muted">${(p.tags ?? []).map(t => `<span style="background:#f3f4f6;padding:1px 6px;border-radius:4px;font-size:11px;">${esc(t)}</span>`).join(' ')}</td>
        <td class="muted">${fmtDate(p.createdAt)}</td>
        <td class="muted">${fmtDate(p.updatedAt)}</td>
        <td style="text-align:right;font-variant-numeric:tabular-nums">${invoiceTotal}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  function computeInvoiceTotal(project) {
    if (!project.invoices || project.invoices.length === 0) return '<span class="muted">—</span>';
    // Sum all invoice totalWithTax values; group by currency
    const byCurrency = {};
    project.invoices.forEach(inv => {
      const currency = inv.currency || '';
      const amount = inv.totalWithTax ?? inv.total ?? 0;
      byCurrency[currency] = (byCurrency[currency] ?? 0) + amount;
    });
    return Object.entries(byCurrency)
      .map(([cur, amt]) => `${fmtNum(amt)} ${esc(cur)}`)
      .join('<br>');
  }

  function statusBadge(s) {
    const label = STATUS_LABEL[s] ?? `Statut ${s}`;
    return `<span class="status-badge status-${s}">${label}</span>`;
  }

  function fmtDate(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleDateString(undefined, { year:'numeric', month:'short', day:'numeric' }); }
    catch { return iso; }
  }

  function fmtNum(n) {
    if (n == null) return '—';
    return new Intl.NumberFormat(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
  }

  function esc(str) {
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  // ── Helpers ───────────────────────────────────────────────────────────────

  function showProgress(visible, pct = 0, label = '') {
    const wrap = document.getElementById('progressWrap');
    wrap.style.display = visible ? '' : 'none';
    document.getElementById('progressFill').style.width = `${pct}%`;
    document.getElementById('progressLabel').textContent = label;
  }

  function clearFilters() {
    ['fStatus','fCustomer','fSubdivision','fInstance','fReference','fTag','fDateFrom','fDateTo'].forEach(id => {
      const el = document.getElementById(id);
      el.tagName === 'SELECT' ? (el.selectedIndex = 0) : (el.value = '');
    });
  }

  // ── Detail panel ─────────────────────────────────────────────────────────

  let activeRow = null;

  async function openPanel(project) {
    // Highlight row
    if (activeRow) activeRow.classList.remove('active');
    activeRow = document.querySelector(`tr[data-guid="${project.guid}"]`);
    if (activeRow) activeRow.classList.add('active');

    // Open panel with basic info immediately
    document.getElementById('panelTitle').textContent = project.name || '—';
    document.getElementById('panelSub').textContent = [project.number, project.reference].filter(Boolean).join(' · ');
    document.getElementById('panelBody').innerHTML = '<div class="panel-loading"><div class="spinner"></div> Chargement des détails…</div>';
    document.getElementById('panelOverlay').classList.add('open');
    document.getElementById('sidePanel').classList.add('open');

    try {
      const res = await fetch(`${BASE}/projects/${project.guid}`, {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const p = await res.json();
      renderPanel(p);
    } catch (e) {
      document.getElementById('panelBody').innerHTML = `<div class="alert alert-error">Erreur : ${esc(e.message)}</div>`;
    }
  }

  function closePanel() {
    document.getElementById('panelOverlay').classList.remove('open');
    document.getElementById('sidePanel').classList.remove('open');
    if (activeRow) { activeRow.classList.remove('active'); activeRow = null; }
  }

  document.addEventListener('keydown', e => { if (e.key === 'Escape') closePanel(); });

  function renderPanel(p) {
    let html = '';

    // Infos générales
    html += `<div class="detail-section">
      <h4>Informations générales</h4>
      <div class="detail-grid">
        <div class="detail-item"><div class="di-label">Numéro</div><div class="di-value">${esc(p.number ?? '—')}</div></div>
        <div class="detail-item"><div class="di-label">Référence</div><div class="di-value">${esc(p.reference ?? '—')}</div></div>
        <div class="detail-item"><div class="di-label">Statut</div><div class="di-value">${statusBadge(p.status)}</div></div>
        <div class="detail-item"><div class="di-label">Subdivision</div><div class="di-value">${esc(p.subdivision ?? '—')}</div></div>
        <div class="detail-item"><div class="di-label">Instance</div><div class="di-value">${esc(p.instance ?? '—')}</div></div>
        <div class="detail-item"><div class="di-label">Créé le</div><div class="di-value">${fmtDate(p.createdAt)}</div></div>
        <div class="detail-item"><div class="di-label">Modifié le</div><div class="di-value">${fmtDate(p.updatedAt)}</div></div>
      </div>
    </div>`;

    // Client
    if (p.customer) {
      const c = p.customer;
      html += `<div class="detail-section">
        <h4>Client</h4>
        <div class="detail-grid">
          <div class="detail-item"><div class="di-label">Nom</div><div class="di-value">${esc(c.name ?? '—')}</div></div>
          <div class="detail-item"><div class="di-label">Email</div><div class="di-value">${esc(c.email ?? '—')}</div></div>
          <div class="detail-item"><div class="di-label">Téléphone</div><div class="di-value">${esc(c.phone ?? '—')}</div></div>
          <div class="detail-item"><div class="di-label">Description</div><div class="di-value">${esc(c.description ?? '—')}</div></div>
        </div>
      </div>`;
    }

    // Assigné à
    if (p.assignedTo) {
      const a = p.assignedTo;
      html += `<div class="detail-section">
        <h4>Assigné à</h4>
        <div class="detail-grid">
          <div class="detail-item"><div class="di-label">Nom</div><div class="di-value">${esc(a.name ?? '—')}</div></div>
          <div class="detail-item"><div class="di-label">Email</div><div class="di-value">${esc(a.email ?? '—')}</div></div>
          <div class="detail-item"><div class="di-label">Téléphone</div><div class="di-value">${esc(a.phone ?? '—')}</div></div>
        </div>
      </div>`;
    }

    // Étiquettes
    if (p.tags && p.tags.length > 0) {
      html += `<div class="detail-section">
        <h4>Étiquettes</h4>
        <div>${p.tags.map(t => `<span class="tag-pill">${esc(t)}</span>`).join('')}</div>
      </div>`;
    }

    // Factures
    if (p.invoices && p.invoices.length > 0) {
      html += `<div class="detail-section"><h4>Factures (${p.invoices.length})</h4>`;
      p.invoices.forEach((inv, i) => {
        const typeLabel = inv.type != null ? `Type ${inv.type}` : '';
        html += `<div class="invoice-block">
          <div class="invoice-block-header">
            <span>Facture ${i + 1}${typeLabel ? ' · ' + typeLabel : ''}${inv.currency ? ' · ' + esc(inv.currency) : ''}</span>
            <span>${fmtNum(inv.totalWithTax ?? inv.total)} ${esc(inv.currency ?? '')}</span>
          </div>`;

        if (inv.lines && inv.lines.length > 0) {
          html += `<div class="invoice-table-wrap"><table class="invoice-lines">
            <thead><tr>
              <th>Désignation</th>
              <th>Catégorie</th>
              <th>Réf. (SKU)</th>
              <th style="text-align:right">Qté</th>
              <th style="text-align:right">P.U.</th>
              <th style="text-align:right">Montant</th>
              <th style="text-align:right">Remise %</th>
              <th style="text-align:right">Remise</th>
              <th style="text-align:right">Total HT</th>
              <th style="text-align:right">TVA %</th>
              <th style="text-align:right">TVA</th>
              <th style="text-align:right">Total TTC</th>
            </tr></thead><tbody>`;
          inv.lines.forEach(l => {
            html += `<tr>
              <td>${esc(l.name ?? '')}<br><span style="font-size:11px;color:#9ca3af">${esc(l.contextName ?? '')}</span></td>
              <td class="muted">${esc(l.categoryName ?? l.categoryCode ?? '')}</td>
              <td class="muted">${esc(l.sku ?? '')}</td>
              <td style="text-align:right">${l.quantity ?? ''}</td>
              <td style="text-align:right">${fmtNum(l.unitPrice)}</td>
              <td style="text-align:right">${fmtNum(l.amount)}</td>
              <td style="text-align:right">${l.discountRate != null ? l.discountRate + ' %' : '—'}</td>
              <td style="text-align:right">${fmtNum(l.discountAmount)}</td>
              <td style="text-align:right">${fmtNum(l.total)}</td>
              <td style="text-align:right">${l.taxRate != null ? l.taxRate + ' %' : '—'}</td>
              <td style="text-align:right">${fmtNum(l.taxAmount)}</td>
              <td style="text-align:right"><strong>${fmtNum(l.totalWithTax)}</strong></td>
            </tr>`;
          });
          html += `</tbody></table></div>`;
        }

        html += `<div class="invoice-totals">
          <div class="total-row"><span>Sous-total</span><span>${fmtNum(inv.subtotal)} ${esc(inv.currency ?? '')}</span></div>`;
        if (inv.discount) html += `<div class="total-row"><span>Remise</span><span>-${fmtNum(inv.discount)} ${esc(inv.currency ?? '')}</span></div>`;
        if (inv.tax)      html += `<div class="total-row"><span>TVA</span><span>${fmtNum(inv.tax)} ${esc(inv.currency ?? '')}</span></div>`;
        html += `<div class="total-row"><strong>Total TTC</strong><strong>${fmtNum(inv.totalWithTax ?? inv.total)} ${esc(inv.currency ?? '')}</strong></div>
        </div></div>`;
      });
      html += `</div>`;
    }

    // Champs personnalisés
    if (p.customFields) {
      try {
        const cf = typeof p.customFields === 'string' ? JSON.parse(p.customFields) : p.customFields;
        const entries = Object.entries(cf).filter(([,v]) => v !== null && v !== '');
        if (entries.length > 0) {
          html += `<div class="detail-section"><h4>Champs personnalisés</h4><div class="detail-grid">`;
          entries.forEach(([k, v]) => {
            html += `<div class="detail-item"><div class="di-label">${esc(k)}</div><div class="di-value">${esc(String(v))}</div></div>`;
          });
          html += `</div></div>`;
        }
      } catch {}
    }

    // Fichiers
    if (p.files && p.files.length > 0) {
      html += `<div class="detail-section"><h4>Fichiers (${p.files.length})</h4>`;
      p.files.forEach(f => {
        html += `<div style="font-size:13px;margin-bottom:4px">📎 ${f.url ? `<a href="${esc(f.url)}" target="_blank" rel="noopener">${esc(f.name ?? f.url)}</a>` : esc(f.name ?? '—')}</div>`;
      });
      html += `</div>`;
    }

    document.getElementById('panelBody').innerHTML = html;
  }

  // ── Export Excel ──────────────────────────────────────────────────────────

  function downloadExcel() {
    if (!currentItems.length) return;

    const rows = currentItems.map(p => {
      const invoiceTotals = {};
      (p.invoices ?? []).forEach(inv => {
        const cur = inv.currency || '';
        invoiceTotals[cur] = (invoiceTotals[cur] ?? 0) + (inv.totalWithTax ?? inv.total ?? 0);
      });
      const totalStr = Object.entries(invoiceTotals)
        .map(([cur, amt]) => `${amt.toFixed(2)} ${cur}`)
        .join(' | ');

      return {
        'N°':            p.number ?? '',
        'Référence':     p.reference ?? '',
        'Nom':           p.name ?? '',
        'Statut':        STATUS_LABEL[p.status] ?? `Statut ${p.status}`,
        'Client':        p.customer?.name ?? p.customer?.identity ?? '',
        'Subdivision':   p.subdivision ?? '',
        'Étiquettes':    (p.tags ?? []).join(', '),
        'Créé le':       p.createdAt ? new Date(p.createdAt).toLocaleDateString('fr-FR') : '',
        'Modifié le':    p.updatedAt ? new Date(p.updatedAt).toLocaleDateString('fr-FR') : '',
        'Total facturé': totalStr,
      };
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Projets');

    // Auto column width
    const cols = Object.keys(rows[0]).map(key => ({
      wch: Math.max(key.length, ...rows.map(r => String(r[key] ?? '').length))
    }));
    ws['!cols'] = cols;

    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `prodboard_projets_${date}.xlsx`);
  }

  // ── Tabs ──────────────────────────────────────────────────────────────────

  function switchTab(name) {
    document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('tab-' + name).classList.add('active');
    document.querySelector(`.tab-btn[onclick="switchTab('${name}')"]`).classList.add('active');
  }

  // Allow Enter key in auth fields to trigger connect
  ['company','privateKey'].forEach(id => {
    document.getElementById(id).addEventListener('keydown', e => { if (e.key === 'Enter') connect(); });
  });
</script>

</body>
</html>
