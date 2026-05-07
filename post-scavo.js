// ══════════════════════════════════════
// PDF.js WORKER
// ══════════════════════════════════════

if (typeof pdfjsLib !== 'undefined') {
  try {
    const workerCode = "importScripts('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js');";
    const blob = new Blob([workerCode], { type: 'application/javascript' });
    pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
  } catch(e) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = '';
  }
}

// ══════════════════════════════════════
// CONFIGURAZIONE
// ══════════════════════════════════════

const SP_SITE_URL         = "https://akhet.sharepoint.com/sites/Akhet-Postscavo";
const SP_CANTIERI_SITE_URL = "https://akhet.sharepoint.com/sites/Akhet-Giornaledeilavori";
const FOTO_DRIVE_NAME = "Foto";
const CATALOGO_LIST   = "CatalogoFoto";
const CANTIERI_LIST   = "Cantieri";
const UTENTI_LIST     = "Utenti";
const VISTE           = ["N","NE","E","SE","S","SW","W","NW"];
// Secret condiviso per magic link (deve essere uguale in index.html)
const FIRMA_SECRET    = "AkhetPostScavo2025";

const MSAL_CONFIG = {
  auth: {
    clientId: "1f1379c3-6946-4ef1-8977-35d32435d409",
    authority: "https://login.microsoftonline.com/a1b7202e-cd9a-4043-98bf-ec0055b06878",
    knownAuthorities: ["login.microsoftonline.com"],
    redirectUri: window.location.origin + window.location.pathname,
    navigateToLoginRequestUrl: false
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
};
const LOGIN_REQUEST = { scopes: ["Sites.ReadWrite.All","User.Read","Files.ReadWrite.All"] };

// ── Stato globale ──────────────────────
let msalInstance        = null;
let currentUser         = null;
let accessToken         = null;
let siteId              = null;
let cantieriSiteId      = null;
let catalogoListId      = null;
let utentiListId        = null;
let currentUserRole     = 'operator'; // 'operator' | 'supervisor' | 'administrative'
let isExternalUser      = false;
let catalogoAggiornato  = false;  // true dopo click "Aggiorna"
let catalogoValidato    = false;  // true dopo click "Valida"
let cantieriData        = [];
let currentProj         = null;
let bulkFiles           = [];
let catalogRows         = [];     // ogni row può avere spItemId se caricata da SP
let rowCounter          = 0;
let uploadedFotoData    = [];
let _pickerRowId        = null;

// ══════════════════════════════════════
// AUTENTICAZIONE
// ══════════════════════════════════════

function getMsal() {
  if (!msalInstance) msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
  return msalInstance;
}

// ── Firma / verifica magic link ────────────────────────
async function firmaPayload(payloadB64) {
  const enc   = new TextEncoder();
  const key   = await crypto.subtle.importKey('raw', enc.encode(FIRMA_SECRET), { name:'HMAC', hash:'SHA-256' }, false, ['sign']);
  const sig   = await crypto.subtle.sign('HMAC', key, enc.encode(payloadB64));
  return btoa(String.fromCharCode(...new Uint8Array(sig)));
}
async function verificaFirma(payloadB64, firma) {
  try {
    const expected = await firmaPayload(payloadB64);
    return expected === firma;
  } catch(e) { return false; }
}

// ── Controllo magic link all'avvio ─────────────────────
async function checkMagicLinkToken() {
  const params    = new URLSearchParams(window.location.search);
  const payloadB64 = params.get('td');
  const firma     = params.get('ts');
  if (!payloadB64 || !firma) return false;

  const errDiv = document.getElementById('loginError');
  errDiv.style.display = 'block';
  errDiv.className     = 'info';
  errDiv.textContent   = '⏳ Verifica link di accesso...';

  const valida = await verificaFirma(payloadB64, decodeURIComponent(firma));
  if (!valida) {
    errDiv.className   = 'msg-error';
    errDiv.textContent = '❌ Link non valido. Contatta il tuo responsabile.';
    return true;
  }

  const payload = JSON.parse(decodeURIComponent(escape(atob(payloadB64))));
  if (Date.now() > payload.exp) {
    errDiv.className   = 'msg-error';
    errDiv.textContent = `❌ Link scaduto il ${new Date(payload.exp).toLocaleDateString('it-IT')}. Chiedi un nuovo link.`;
    return true;
  }

  window.history.replaceState({}, document.title, window.location.pathname);
  isExternalUser  = true;
  currentUser     = { username: payload.email, name: payload.nome, isExternal: true };
  currentUserRole = payload.ruolo || 'operator';
  await initializeAppEsterno(payload);
  return true;
}

async function initializeAppEsterno(payload) {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('app').style.display         = 'block';
  document.getElementById('userName').textContent      = payload.nome || payload.email;
  document.getElementById('userRoleBadge').textContent = `👤 ${payload.ruolo || 'Operator'} (Esterno)`;
  await getSiteId();
  await loadCantieri();
  await ensureCatalogoList();
  updateButtonStates();
}

// ── Genera magic link (solo Supervisor/Admin) ──────────
async function generaMagicLink() {
  const emailInput = document.getElementById('magicLinkTargetEmail');
  const nomeInput  = document.getElementById('magicLinkTargetNome');
  const ruoloInput = document.getElementById('magicLinkTargetRuolo');
  if (!emailInput || !nomeInput) return;
  const email = emailInput.value.trim();
  const nome  = nomeInput.value.trim();
  const ruolo = ruoloInput?.value || 'operator';
  if (!email || !nome) { alert('Inserisci email e nome.'); return; }

  const scadenza = new Date();
  scadenza.setDate(scadenza.getDate() + 30);
  const payload = JSON.stringify({ email, nome, ruolo, exp: scadenza.getTime() });
  const payloadB64 = btoa(unescape(encodeURIComponent(payload)));
  const firma = await firmaPayload(payloadB64);
  const link  = `${window.location.href.split('?')[0]}?td=${payloadB64}&ts=${encodeURIComponent(firma)}`;

  document.getElementById('magicLinkResult').value = link;
  document.getElementById('magicLinkBox').style.display = 'block';
}

async function loginMicrosoft() {
  const btn    = document.getElementById('btnLogin');
  const errDiv = document.getElementById('loginError');
  try {
    btn.disabled = true; btn.textContent = '⏳ Connessione...';
    errDiv.style.display = 'none';
    const resp  = await getMsal().loginPopup(LOGIN_REQUEST);
    currentUser = resp.account;
    const token = await getMsal().acquireTokenSilent({ scopes: LOGIN_REQUEST.scopes, account: currentUser });
    accessToken = token.accessToken;
    await initializeApp();
  } catch(e) {
    errDiv.textContent = `Errore: ${e.message}`;
    errDiv.style.display = 'block';
    btn.disabled = false; btn.textContent = '🔐 Accedi con Microsoft 365';
  }
}

function logout() {
  sessionStorage.clear();
  getMsal().logoutPopup().then(() => window.location.reload()).catch(() => window.location.reload());
}
function clearSessionAndReload() { sessionStorage.clear(); localStorage.clear(); window.location.reload(); }

// Bottone "Accedi con Link Email" — spiega all'utente cosa fare
function richiediAccesso() {
  const email    = document.getElementById('magicLinkEmail').value.trim();
  const errDiv   = document.getElementById('loginError');
  if (!email) {
    errDiv.style.display    = 'block';
    errDiv.style.background = '#fff3cd';
    errDiv.style.color      = '#856404';
    errDiv.textContent      = '⚠️ Inserisci la tua email per continuare.';
    return;
  }
  errDiv.style.display    = 'block';
  errDiv.style.background = '#e8f5e9';
  errDiv.style.color      = '#2e7d32';
  errDiv.textContent      = `📧 Chiedi al tuo responsabile di inviarti il link di accesso per ${email}. Quando ricevi il link, aprilo direttamente dal browser.`;
}

// ══════════════════════════════════════
// SHAREPOINT API
// ══════════════════════════════════════

async function spFetch(url, opts = {}) {
  const res = await fetch(url, {
    ...opts,
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json', ...(opts.headers || {}) }
  });
  if (!res.ok) { const t = await res.text().catch(()=>''); throw new Error(`SP ${res.status}: ${t.slice(0,200)}`); }
  return res.json();
}

// Legge tutti i file di una cartella su Drive seguendo la paginazione (nessun limite)
async function getDriveChildren(driveId, folderPath, selectFields) {
  const sel   = selectFields ? `&$select=${selectFields}` : '';
  let   url   = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${encodeURIComponent(folderPath)}:/children?$top=200${sel}`;
  const items = [];
  while (url) {
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!res.ok) break;
    const d = await res.json();
    items.push(...(d.value || []));
    url = d['@odata.nextLink'] || null;
  }
  return items;
}

async function getSiteId() {
  const u1 = new URL(SP_SITE_URL);
  const d1 = await spFetch(`https://graph.microsoft.com/v1.0/sites/${u1.hostname}:${u1.pathname}`);
  siteId = d1.id;
  const u2 = new URL(SP_CANTIERI_SITE_URL);
  const d2 = await spFetch(`https://graph.microsoft.com/v1.0/sites/${u2.hostname}:${u2.pathname}`);
  cantieriSiteId = d2.id;
}

async function loadCurrentUserRole() {
  try {
    // Recupera la lista Utenti su Akhet-Giornale dei lavori
    const lists = await spFetch(`https://graph.microsoft.com/v1.0/sites/${cantieriSiteId}/lists?$filter=displayName eq '${UTENTI_LIST}'`);
    if (!lists.value.length) { console.warn('Lista Utenti non trovata'); return; }
    utentiListId = lists.value[0].id;

    // Email dell'utente loggato
    const email = (currentUser?.username || currentUser?.mail || '').toLowerCase();
    const items = await spFetch(`https://graph.microsoft.com/v1.0/sites/${cantieriSiteId}/lists/${utentiListId}/items?$expand=fields&$top=999`);

    const utente = items.value.find(i => {
      const e = (i.fields.Email || i.fields.Title || '').toLowerCase();
      // Gestisce anche guest M365 tipo mario.rossi_domain.com#EXT#@tenant
      const emailPulita = email.includes('#ext#') ? email.split('#ext#')[0].replace(/_/g, '@').replace(/@.*$/, (m, o, s) => '@' + s.split('@').pop()) : email;
      return e === email || e === emailPulita;
    });

    if (utente) {
      currentUserRole = (utente.fields.Ruolo || 'operator').toLowerCase();
      console.log(`👤 Ruolo: ${currentUserRole}`);
    } else {
      console.warn('Utente non trovato nella lista Utenti, ruolo default: operator');
      currentUserRole = 'operator';
    }
  } catch(e) {
    console.warn('loadCurrentUserRole errore:', e.message);
    currentUserRole = 'operator';
  }
}

async function loadCantieri() {
  const d = await spFetch(`https://graph.microsoft.com/v1.0/sites/${cantieriSiteId}/lists?$filter=displayName eq '${CANTIERI_LIST}'`);
  if (!d.value.length) throw new Error("Lista 'Cantieri' non trovata su SharePoint.");
  const listId = d.value[0].id;
  const items  = await spFetch(`https://graph.microsoft.com/v1.0/sites/${cantieriSiteId}/lists/${listId}/items?$expand=fields&$top=999`);

  cantieriData = items.value.map(i => ({
    id:                  i.id,
    title:               i.fields.Title || '',
    comune:              i.fields.Comune || '',
    committente:         i.fields.Committente || '',
    codiceProgetto:      i.fields.CodProg || '',
    codiceCommessa:      i.fields.CodComm || '',
    codiceSito:          i.fields.CodiceSito || '',
    responsabileEnte:    i.fields.ResponsabileEnte || '',
    descrizioneProgetto: i.fields.Descrizione || i.fields.Description || i.fields.Descrizione0 || '',
  }));

  // Debug: stampa i campi del primo cantiere per verificare i nomi
  if (items.value.length > 0) {
    console.log('📋 Campi cantiere (primo elemento):', Object.keys(items.value[0].fields).join(', '));
  }

  const sel = document.getElementById('selCantiere');
  sel.innerHTML = '<option value="">— Tutti i cantieri —</option>';
  cantieriData.forEach(c => {
    const o = document.createElement('option');
    o.value = c.id;
    o.textContent = c.title + (c.committente ? ` — ${c.committente}` : '');
    sel.appendChild(o);
  });

  // Popola anche il dropdown della sezione Elenco Tavole con gli stessi dati
  const tavSel = document.getElementById('tavSelCantiere');
  if (tavSel) tavSel.innerHTML = sel.innerHTML;

  // Carica conteggi foto e renderizza griglia cantieri
  await renderCantieriGrid(cantieriData);
}

async function ensureCatalogoList() {
  const d = await spFetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=displayName eq '${CATALOGO_LIST}'`);
  if (!d.value.length) throw new Error(`Lista "${CATALOGO_LIST}" non trovata su SharePoint. Creala prima di procedere.`);
  catalogoListId = d.value[0].id;

  // Stampa in console i nomi interni reali di tutte le colonne (utile per debug)
  try {
    const cols = await spFetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${catalogoListId}/columns`);
    const nomi = cols.value
      .filter(c => !c.readOnly && !c.hidden)
      .map(c => `"${c.name}" (display: "${c.displayName}")`);
    console.log('📋 Colonne CatalogoFoto:', nomi.join('\n'));
  } catch(e) { console.warn('Impossibile leggere colonne:', e.message); }
}

async function getDriveForFoto() {
  const d = await spFetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`);
  console.log('📁 Raccolte disponibili:', d.value.map(x => `"${x.name}"`).join(', '));
  const drive = d.value.find(x => x.name === FOTO_DRIVE_NAME);
  if (!drive) throw new Error(`Raccolta documenti "${FOTO_DRIVE_NAME}" non trovata su Akhet-Postscavo. Creala prima di caricare foto.`);
  return drive;
}

async function uploadFileToDrive(driveId, folderName, filename, blob, retries = 3) {
  const path = `${folderName}/${filename}`;
  for (let attempt = 1; attempt <= retries; attempt++) {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${encodeURIComponent(path)}:/content`,
      { method: 'PUT', headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'image/jpeg' }, body: blob }
    );
    if (res.ok) return res.json();
    if (attempt < retries) {
      console.warn(`⚠️ Upload ${filename} fallito (tentativo ${attempt}/${retries}), riprovo...`);
      await sleep(1500 * attempt);
    } else {
      throw new Error(`Upload ${filename} fallito dopo ${retries} tentativi (${res.status})`);
    }
  }
}

async function saveCatalogoItem(row, idx) {
  const fields = {
    Title:          row.filename,
    NFoto:          row.filename,
    FilenameFoto:   row.filename,
    CantiereTitolo: currentProj?.title          || '',
    CodiceSito:     currentProj?.codiceSito     || '',
    Committente:    currentProj?.committente    || '',
    CodiceProgetto: currentProj?.codiceProgetto || '',
    Comune:         currentProj?.comune         || '',
    Operatore:      row.operatore   || '',
    Descrizione:    row.descrizione || '',
    VistaDa:        row.vista       || '',
    Contesto:       row.contesto    || '',
  };

  // DataFoto: invia solo se valorizzata e nel formato ISO corretto
  if (row.data) {
    fields['DataFoto'] = row.data.includes('T') ? row.data : row.data + 'T00:00:00Z';
  }

  // Campi con caratteri speciali nel nome interno
  fields['Localit_x00e0_'] = row.localita || '';
  fields['Qualit_x00e0_']  = row.hq ? 'Alta qualità' : 'Compressa';

  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${catalogoListId}/items`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields })
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    console.error('saveCatalogoItem errore body inviato:', JSON.stringify(fields));
    console.error('saveCatalogoItem errore risposta:', txt);
    throw new Error(`Salvataggio fallito (${res.status}): ${txt.slice(0, 300)}`);
  }
  return res.json();
}

async function getExistingPhotoCount() {
  try {
    const drive = await getDriveForFoto();
    const folderName = buildFolderName();
    const files = await getDriveChildren(drive.id, folderName, 'name');
    const imageFiles = files.filter(f => /\.(jpg|jpeg|png|heic)$/i.test(f.name));
    console.log(`📊 Foto esistenti nella cartella "${folderName}": ${imageFiles.length}`);
    return imageFiles.length;
  } catch(e) {
    console.warn('getExistingPhotoCount errore:', e.message);
    return 0;
  }
}

async function updateCatalogoItem(row) {
  if (!row.spItemId) return;
  const fields = {
    Title:          row.filename,
    NFoto:          row.filename,
    FilenameFoto:   row.filename,
    CantiereTitolo: currentProj?.title          || '',
    CodiceSito:     currentProj?.codiceSito     || '',
    Committente:    currentProj?.committente    || '',
    CodiceProgetto: currentProj?.codiceProgetto || '',
    Comune:         currentProj?.comune         || '',
    Operatore:      row.operatore   || '',
    Descrizione:    row.descrizione || '',
    VistaDa:        row.vista       || '',
    Contesto:       row.contesto    || '',
  };
  if (row.data) fields['DataFoto'] = row.data.includes('T') ? row.data : row.data + 'T00:00:00Z';
  fields['Localit_x00e0_'] = row.localita || '';
  fields['Qualit_x00e0_']  = row.hq ? 'Alta qualità' : 'Compressa';

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${catalogoListId}/items/${row.spItemId}/fields`,
    { method: 'PATCH', headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' }, body: JSON.stringify(fields) }
  );
  if (!res.ok) throw new Error(`Aggiornamento item ${row.spItemId} fallito (${res.status})`);
}

async function getExistingCatalogItems() {
  if (!catalogoListId) return [];
  let allItems = [], nextLink = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${catalogoListId}/items?$expand=fields&$top=999`;
  while (nextLink) {
    const d = await spFetch(nextLink);
    allItems = allItems.concat(d.value || []);
    nextLink = d['@odata.nextLink'] || null;
  }
  return allItems
    .filter(i => i.fields?.CantiereTitolo === currentProj?.title)
    .sort((a, b) => {
      const da = a.fields?.DataFoto ? new Date(a.fields.DataFoto) : new Date(0);
      const db = b.fields?.DataFoto ? new Date(b.fields.DataFoto) : new Date(0);
      return da - db;
    })
    .map(i => ({
      spItemId:    i.id,
      filename:    i.fields.NFoto || i.fields.Title || '',
      operatore:   i.fields.Operatore || '',
      data:        i.fields.DataFoto ? i.fields.DataFoto.split('T')[0] : '',
      localita:    i.fields['Localit_x00e0_'] || '',
      descrizione: i.fields.Descrizione || '',
      vista:       i.fields.VistaDa || '',
      contesto:    i.fields.Contesto || '',
      hq:          i.fields['Qualit_x00e0_'] === 'Alta qualità',
      previewUrl:  '',  // non disponibile da SP
      id:          ++rowCounter,
      exif:        {},
      file:        null,
    }));
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ══════════════════════════════════════
// INIT
// ══════════════════════════════════════

async function initializeApp() {
  const overlay = document.getElementById('loadingOverlay');
  const setStep = msg => { const el = document.getElementById('loadingStep'); if(el) el.textContent = msg; };
  try {
    document.getElementById('loginScreen').style.display = 'none';
    document.getElementById('app').style.display = 'block';
    overlay.style.display = 'flex';
    document.getElementById('userName').textContent = currentUser.name || currentUser.username;

    setStep('Connessione a SharePoint...');
    await getSiteId();

    setStep('Caricamento ruolo utente...');
    await loadCurrentUserRole();
    const roleLabels = { operator: '👷 Operator', supervisor: '🔑 Supervisor', administrative: '⚙️ Administrative' };
    document.getElementById('userRoleBadge').textContent = roleLabels[currentUserRole] || '👷 Operator';

    setStep('Caricamento cantieri...');
    await loadCantieri();

    setStep('Verifica lista catalogo...');
    await ensureCatalogoList();

    overlay.style.display = 'none';
    updateButtonStates();
  } catch(e) {
    overlay.style.display = 'none';
    showMsg('❌ ' + e.message, 'error');
  }
}

// Mostra/nasconde bottoni in base al ruolo e allo stato
function updateButtonStates() {
  const isSupervisor = ['supervisor','administrative'].includes(currentUserRole);
  const hasCatalog   = catalogRows.length > 0;

  // Bottone Valida — solo Supervisor/Admin, disponibile appena ci sono righe
  const btnValida = document.getElementById('btnValida');
  if (btnValida) {
    btnValida.style.display = isSupervisor ? 'inline-block' : 'none';
    btnValida.disabled      = !hasCatalog;
    btnValida.style.opacity = hasCatalog ? '1' : '0.5';
  }

  // Bottone Genera catalogo — solo Supervisor/Admin, solo dopo Valida
  const btnGenera = document.getElementById('btnGeneraCatalogo');
  if (btnGenera) {
    btnGenera.style.display = isSupervisor ? 'inline-block' : 'none';
    btnGenera.disabled      = !catalogoValidato;
    btnGenera.style.opacity = catalogoValidato ? '1' : '0.5';
  }

  // Badge ruolo nella griglia cantieri
  const mgLink = document.getElementById('btnGeneraMagicLink');
  if (mgLink) mgLink.style.display = isSupervisor ? 'inline-block' : 'none';
}

// ══════════════════════════════════════
// AVVIO — magic link + ripristino sessione MSAL
// ══════════════════════════════════════

(async () => {
  // Prima controlla magic link nell'URL
  const handled = await checkMagicLinkToken().catch(() => false);
  if (handled) return;

  // Poi prova ripristino sessione MSAL
  try {
    const accounts = getMsal().getAllAccounts();
    if (accounts.length > 0) {
      currentUser = accounts[0];
      const token = await getMsal().acquireTokenSilent({ scopes: LOGIN_REQUEST.scopes, account: currentUser });
      accessToken = token.accessToken;
      await initializeApp();
    }
  } catch(e) { console.log('Nessuna sessione attiva:', e.message); }
})();

// ══════════════════════════════════════
// FASE 1 — Selezione cantiere
// ══════════════════════════════════════

function onCantiereChange() {
  const id = document.getElementById('selCantiere').value;
  if (id) {
    selectCantiere(id);
  } else {
    currentProj = null;
    document.getElementById('cantInfoBar').style.display = 'none';
    document.getElementById('modeSelector').style.display = 'none';
    // De-seleziona le card
    document.querySelectorAll('.cantiere-card').forEach(c => c.classList.remove('selected'));
  }
}

function selectCantiere(id) {
  currentProj = cantieriData.find(c => c.id === id) || null;
  if (!currentProj) return;

  // Aggiorna tendina
  document.getElementById('selCantiere').value = id;

  // Evidenzia card
  document.querySelectorAll('.cantiere-card').forEach(c => {
    c.classList.toggle('selected', c.dataset.id === id);
  });

  // Info bar
  const bar = document.getElementById('cantInfoBar');
  bar.style.display = 'block';
  bar.innerHTML = `<strong>${currentProj.title}</strong>`
    + (currentProj.committente   ? ` &nbsp;|&nbsp; ${currentProj.committente}` : '')
    + (currentProj.codiceSito    ? ` &nbsp;|&nbsp; Cod. sito: <strong>${currentProj.codiceSito}</strong>` : '')
    + (currentProj.codiceProgetto ? ` &nbsp;|&nbsp; Prog. <strong>${currentProj.codiceProgetto}</strong>` : '');
  document.getElementById('modeSelector').style.display = 'block';
  window.scrollTo(0, document.getElementById('cantInfoBar').offsetTop - 20);
}

function filterCantieri() {
  const q = document.getElementById('searchCantiere').value.toLowerCase().trim();
  const filtered = q
    ? cantieriData.filter(c =>
        (c.title       || '').toLowerCase().includes(q) ||
        (c.comune      || '').toLowerCase().includes(q) ||
        (c.committente || '').toLowerCase().includes(q) ||
        (c.codiceSito  || '').toLowerCase().includes(q))
    : cantieriData;

  // Aggiorna tendina
  const sel = document.getElementById('selCantiere');
  sel.innerHTML = '<option value="">— Tutti i cantieri —</option>';
  filtered.forEach(c => {
    const o = document.createElement('option');
    o.value = c.id;
    o.textContent = c.title + (c.committente ? ` — ${c.committente}` : '');
    sel.appendChild(o);
  });

  // Aggiorna griglia (senza ricaricare i conteggi)
  buildCantieriCards(filtered);
}

async function renderCantieriGrid(list) {
  // Prima renderizza le card senza conteggi per mostrare subito qualcosa
  buildCantieriCards(list);

  // Poi tenta di caricare i conteggi foto da SharePoint in background
  if (!catalogoListId) return;
  try {
    // Recupera tutti gli item espandendo i fields (senza $select annidato che causa problemi)
    let allItems = [];
    let nextLink = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${catalogoListId}/items?$expand=fields&$top=999`;
    while (nextLink) {
      const d = await spFetch(nextLink);
      allItems = allItems.concat(d.value || []);
      nextLink = d['@odata.nextLink'] || null;
    }

    // Conta per cantiere usando CantiereTitolo
    const counts = {};
    allItems.forEach(item => {
      const t = item.fields?.CantiereTitolo || '';
      if (t) counts[t] = (counts[t] || 0) + 1;
    });

    console.log('📊 Conteggi foto per cantiere:', counts);
    // Aggiorna badge nelle card
    list.forEach(c => {
      const badge = document.getElementById(`badge-${c.id}`);
      if (badge) {
        const n = counts[c.title] || 0;
        badge.textContent = n ? `${n} foto` : 'Nessuna foto';
        badge.style.background = n ? '#b8423a' : '#aaa';
      }
    });
  } catch(e) { console.log('Conteggi foto non disponibili:', e.message); }
}

function buildCantieriCards(list) {
  const grid = document.getElementById('cantieriGrid');
  if (!grid) return;
  grid.innerHTML = '';
  list.forEach(c => {
    const card = document.createElement('div');
    card.className = 'cantiere-card' + (currentProj?.id === c.id ? ' selected' : '');
    card.dataset.id = c.id;
    card.onclick = () => selectCantiere(c.id);
    card.innerHTML = `
      <h4>${esc(c.title)}</h4>
      <p>${esc(c.comune || '')}${c.committente ? ` · ${esc(c.committente)}` : ''}</p>
      ${c.codiceProgetto ? `<p style="font-size:11px;color:#aaa">Prog. ${esc(c.codiceProgetto)}</p>` : ''}
      <span class="foto-count-badge" id="badge-${c.id}" style="background:#aaa">...</span>`;
    grid.appendChild(card);
  });
}

// Torna alla fase 2 senza perdere il cantiere né le righe già compilate
function tornaAdAggiungereFoto() {
  bulkFiles = [];
  document.getElementById('previewGrid').innerHTML = '';
  document.getElementById('uploadActions').style.display = 'none';
  document.getElementById('uploadLegend').style.display  = 'none';
  document.getElementById('progressBar').style.width = '0%';
  document.getElementById('uploadLog').innerHTML    = '';
  const btn = document.getElementById('btnStartUpload');
  if (btn) btn.disabled = false;
  syncInfoBars();
  goToPhase('phase-upload', 2);
}

// Torna ai cantieri resettando tutto
function tornaIndietro() {
  resetAll();
}

// PDF dal catalogo corrente (fase 4), prima del salvataggio
function generaPDFCatalogo() {
  if (!catalogRows.length) { alert('Nessuna foto nel catalogo.'); return; }
  const rows = catalogRows.map(r => ({
    num:         r.filename.replace(/\.jpg$/i, ''),
    operatore:   r.operatore || '',
    dataFmt:     r.data ? new Date(r.data + 'T00:00:00').toLocaleDateString('it-IT') : '',
    localita:    r.localita   || '',
    descrizione: r.descrizione || '',
    vista:       r.vista      || '',
    contesto:    r.contesto   || '',
  }));
  openPrintWindow(buildPdfHtml(rows), 'elenco_foto_bozza');
}

// ══════════════════════════════════════
// NAVIGAZIONE
// ══════════════════════════════════════

function goToPhase(phaseId, stepNum) {
  document.querySelectorAll('.phase').forEach(p => p.classList.remove('active'));
  document.getElementById(phaseId).classList.add('active');
  document.querySelectorAll('.step-indicator .step').forEach((s, i) => {
    s.classList.remove('active','done');
    if (i + 1 < stepNum) s.classList.add('done');
    else if (i + 1 === stepNum) s.classList.add('active');
  });
  window.scrollTo(0, 0);
}

function showMainSection(name) {
  ['foto','us','tavole'].forEach(n => {
    const sec = document.getElementById('sec-' + n);
    const btn = document.getElementById('mnav-' + n);
    if (sec) sec.classList.toggle('active', n === name);
    if (btn) btn.classList.toggle('active', n === name);
  });
  window.scrollTo(0, 0);
}

function syncInfoBars() {
  const html = document.getElementById('cantInfoBar').innerHTML;
  ['cantInfoBar2','cantInfoBar4','cantInfoBar5'].forEach(id => { const el = document.getElementById(id); if(el) el.innerHTML = html; });
}

// ══════════════════════════════════════
// FASE 2 — Bulk upload + qualità
// ══════════════════════════════════════

async function onBulkSelect(event) {
  const files = Array.from(event.target.files);
  if (!files.length) return;
  event.target.value = '';
  const grid = document.getElementById('previewGrid');
  for (const file of files) {
    const id         = ++rowCounter;
    const previewUrl = await fileToDataUrl(file);
    const exif       = await readExif(file);
    bulkFiles.push({ id, file, previewUrl, hq: false, exif });
    grid.appendChild(buildPreviewCard({ id, previewUrl, hq: false, exif }));
  }
  document.getElementById('uploadActions').style.display = 'flex';
  document.getElementById('uploadLegend').style.display  = 'block';
  refreshPreviewNums();
  updateUploadCount();
}

function buildPreviewCard(item) {
  const dateStr = item.exif?.dateStr || '';
  const card    = document.createElement('div');
  card.className = 'foto-preview-card' + (item.hq ? ' hq' : '');
  card.id = 'pcard-' + item.id;
  card.innerHTML = `
    <div class="hq-badge">ALTA Q.</div>
    <img src="${item.previewUrl}" onclick="openLightbox('${item.previewUrl}')">
    <button class="remove-btn" onclick="removePreview(${item.id})">✕</button>
    <div class="card-meta">
      <div class="card-num" id="pnum-${item.id}">—</div>
      ${dateStr ? `<div>${dateStr}</div>` : '<div style="color:#f39c12">📷 no EXIF date</div>'}
    </div>
    <button class="quality-toggle ${item.hq ? 'hq-btn' : 'compressed'}" onclick="toggleQuality(${item.id})">
      ${item.hq ? '⭐ Alta qualità' : '🗜️ Compressa'}
    </button>`;
  return card;
}

function refreshPreviewNums() {
  const sorted = [...bulkFiles].sort(sortByExifDate);
  sorted.forEach((item, idx) => {
    const el = document.getElementById('pnum-' + item.id);
    if (el) {
      const year = item.exif?.dateObj?.getFullYear() || '????';
      el.textContent = buildNumPreview(idx, year);
    }
  });
}

// Versione preview (senza offset globale, che si conosce solo al momento dell'upload)
function buildNumPreview(i, year) {
  const sito = sanitize(currentProj?.codiceSito  || 'SITO');
  const comm = sanitize(currentProj?.committente || 'COMM');
  return `${sito}_${comm}_${year}_${String(i + 1).padStart(3,'0')}`;
}

function toggleQuality(id) {
  const item = bulkFiles.find(f => f.id === id);
  if (!item) return;
  item.hq = !item.hq;
  const card = document.getElementById('pcard-' + id);
  card.className = 'foto-preview-card' + (item.hq ? ' hq' : '');
  const btn = card.querySelector('.quality-toggle');
  btn.className = 'quality-toggle ' + (item.hq ? 'hq-btn' : 'compressed');
  btn.textContent = item.hq ? '⭐ Alta qualità' : '🗜️ Compressa';
  updateUploadCount();
}

function removePreview(id) {
  bulkFiles = bulkFiles.filter(f => f.id !== id);
  document.getElementById('pcard-' + id)?.remove();
  refreshPreviewNums(); updateUploadCount();
  if (!bulkFiles.length) {
    document.getElementById('uploadActions').style.display = 'none';
    document.getElementById('uploadLegend').style.display  = 'none';
  }
}

function selectAllHQ() {
  bulkFiles.forEach(f => { f.hq = true; });
  bulkFiles.forEach(f => {
    const c = document.getElementById('pcard-' + f.id); if (!c) return;
    c.className = 'foto-preview-card hq';
    const b = c.querySelector('.quality-toggle'); b.className = 'quality-toggle hq-btn'; b.textContent = '⭐ Alta qualità';
  });
  updateUploadCount();
}

function selectAllCompressed() {
  bulkFiles.forEach(f => { f.hq = false; });
  bulkFiles.forEach(f => {
    const c = document.getElementById('pcard-' + f.id); if (!c) return;
    c.className = 'foto-preview-card';
    const b = c.querySelector('.quality-toggle'); b.className = 'quality-toggle compressed'; b.textContent = '🗜️ Compressa';
  });
  updateUploadCount();
}

function updateUploadCount() {
  const hq = bulkFiles.filter(f => f.hq).length;
  document.getElementById('uploadCount').textContent =
    `${bulkFiles.length} foto · ${hq} alta qualità · ${bulkFiles.length - hq} compresse`;
}

// ══════════════════════════════════════
// FASE 3 — Upload su SharePoint
// ══════════════════════════════════════

async function startUpload() {
  if (!bulkFiles.length) { alert('Seleziona almeno una foto.'); return; }
  document.getElementById('btnStartUpload').disabled = true;
  goToPhase('phase-uploading', 3);

  const bar     = document.getElementById('progressBar');
  const counter = document.getElementById('uplCounter');
  const msg     = document.getElementById('uplMsg');
  const log     = document.getElementById('uploadLog');
  log.innerHTML = '';
  const addLog  = t => { log.innerHTML += `<div>${t}</div>`; log.scrollTop = log.scrollHeight; };

  // Ordina per data EXIF
  const sorted = [...bulkFiles].sort(sortByExifDate);

  try {
    msg.textContent = 'Connessione al drive SharePoint...';
    const drive   = await getDriveForFoto();
    const driveId = drive.id;

    const folderName = buildFolderName();
    addLog(`📁 Drive: ${drive.name} | Cartella: ${folderName}`);

    // Contatori per anno: max tra file su SP e righe già in catalogRows per quell'anno
    msg.textContent = 'Recupero conteggio foto esistenti per anno...';
    const yearCounters = {};

    // Salva le righe già presenti (da sessione precedente o da rivedi catalogo)
    const existingRows = [...catalogRows];

    for (let i = 0; i < sorted.length; i++) {
      const item = sorted[i];
      const year = item.exif?.dateObj?.getFullYear() || new Date().getFullYear();

      if (yearCounters[year] === undefined) {
        // Conta file su SP
        const spCount = await getExistingPhotoCountForYear(driveId, folderName, year);
        // Conta righe già in catalogRows per questo anno
        const memCount = existingRows.filter(r => {
          if (!r.data) return false;
          return new Date(r.data + 'T00:00:00').getFullYear() === year;
        }).length;
        yearCounters[year] = Math.max(spCount, memCount);
        addLog(`📊 Anno ${year}: ${spCount} file su SP, ${memCount} in memoria → parto da ${yearCounters[year] + 1}`);
      }

      const fname = buildNum(yearCounters[year], year) + '.jpg';

      // Controllo anti-duplicato
      const isDupe = existingRows.some(r => r.filename === fname) ||
                     catalogRows.some(r => r.filename === fname);
      if (isDupe) {
        addLog(`⚠️ Salto ${fname}: nome già presente nel catalogo`);
        yearCounters[year]++;
        continue;
      }

      yearCounters[year]++;

      bar.style.width = Math.round((i / sorted.length) * 100) + '%';
      counter.textContent = `Caricamento ${i + 1} di ${sorted.length}: ${fname}`;

      let blob = item.hq ? item.file : await compImg(item.file);
      await uploadFileToDrive(driveId, folderName, fname, blob);
      addLog(`✅ ${fname} · ${item.hq ? 'originale' : 'compressa'}${item.exif?.dateStr ? ' · ' + item.exif.dateStr : ''}`);

      let localita = '';
      if (item.exif?.lat != null && item.exif?.lng != null) {
        msg.textContent = `Geolocalizzazione ${i + 1}/${sorted.length}...`;
        localita = await reverseGeocode(item.exif.lat, item.exif.lng);
      }

      catalogRows.push({
        id:          item.id,
        file:        item.file,
        previewUrl:  item.previewUrl,
        hq:          item.hq,
        exif:        item.exif,
        filename:    fname,
        spItemId:    null,  // non ancora salvato in CatalogoFoto
        operatore:   currentUser?.name || currentUser?.username || '',
        data:        item.exif?.date || '',
        localita,
        descrizione: '',
        vista:       '',
        contesto:    '',
      });
    }

    bar.style.width = '100%';
    counter.textContent = `✅ ${sorted.length} foto caricate con successo!`;
    addLog('📋 Procedendo alla catalogazione...');
    syncInfoBars();
    renderCatalogRows();
    setTimeout(() => goToPhase('phase-catalog', 4), 700);

  } catch(e) {
    console.error(e);
    goToPhase('phase-upload', 2);
    showMsg('❌ Errore upload: ' + e.message, 'error');
    document.getElementById('btnStartUpload').disabled = false;
  }
}

// ══════════════════════════════════════
// FASE 4 — Catalogazione
// ══════════════════════════════════════

function renderCatalogRows() {
  const tbody = document.getElementById('catalogRows');
  tbody.innerHTML = '';

  // Trova nomi duplicati e avvisa
  const nameCounts = {};
  catalogRows.forEach(r => { if (r.filename) nameCounts[r.filename] = (nameCounts[r.filename] || 0) + 1; });
  const dupeNames = Object.keys(nameCounts).filter(k => nameCounts[k] > 1);
  if (dupeNames.length) {
    showMsg(`⚠️ ${dupeNames.length} nome/i duplicato/i nel catalogo. Clicca "Aggiorna" per rinumerare correttamente.`, 'error');
  }

  catalogRows.forEach((r, i) => {
    const isDupe = nameCounts[r.filename] > 1;

    const tr = document.createElement('tr');
    if (isDupe) tr.style.background = '#fff3cd'; // evidenzia giallo i duplicati

    // Cella thumbnail
    const thumbTd = document.createElement('td');
    thumbTd.className = 'thumb-cell';
    thumbTd.dataset.filename = r.filename;
    const thumbSrc = r.previewUrl || r.thumbUrl || '';
    if (thumbSrc) {
      const img = document.createElement('img');
      img.src = thumbSrc;
      img.style.cssText = 'width:56px;height:56px;object-fit:cover;border-radius:6px;border:1px solid #ddd;cursor:zoom-in';
      img.addEventListener('click', () => openLightboxFull(r));
      thumbTd.appendChild(img);
    } else {
      const ph = document.createElement('div');
      ph.style.cssText = 'width:56px;height:56px;border-radius:6px;border:1px solid #ddd;background:#f5f0e8;display:flex;align-items:center;justify-content:center;font-size:22px';
      ph.textContent = '🖼️';
      thumbTd.appendChild(ph);
    }
    const changeBtn = document.createElement('button');
    changeBtn.className = 'change-btn';
    changeBtn.textContent = '🔄 cambia';
    changeBtn.onclick = () => openSinglePicker(r.id);
    thumbTd.appendChild(changeBtn);
    tr.appendChild(thumbTd);

    tr.insertAdjacentHTML('beforeend', `
      <td class="num-cell">${esc(r.filename.replace(/\.jpg$/i,''))}</td>
      <td style="font-size:11px;white-space:nowrap">
        ${r.hq ? '<span style="color:#b8423a;font-weight:700">⭐ Alta</span>' : '<span style="color:#888">🗜️ Comp.</span>'}
      </td>
      <td><input type="text" class="sm" value="${esc(r.operatore)}" placeholder="Operatore" oninput="updateCatalogField(${r.id},'operatore',this.value)"></td>
      <td>
        <input type="date" class="sm" value="${r.data || ''}" onchange="updateCatalogField(${r.id},'data',this.value)">
        ${r.exif?.date ? '<div class="exif-tag">📷 EXIF</div>' : ''}
      </td>
      <td>
        <input type="text" class="sm" value="${esc(r.localita)}" placeholder="Via/Località" oninput="updateCatalogField(${r.id},'localita',this.value)" style="min-width:135px">
        ${r.exif?.lat != null ? '<div class="exif-tag">📍 GPS</div>' : ''}
      </td>
      <td><input type="text" class="sm" value="${esc(r.descrizione)}" placeholder="Descrizione" oninput="updateCatalogField(${r.id},'descrizione',this.value)" style="min-width:150px"></td>
      <td>
        <select class="sm" onchange="updateCatalogField(${r.id},'vista',this.value)">
          <option value="">—</option>
          ${VISTE.map(v => `<option${v===r.vista?' selected':''}>${v}</option>`).join('')}
        </select>
      </td>
      <td><input type="text" list="contesti-list" class="contesto-input" value="${esc(r.contesto)}" placeholder="Contesto…" oninput="updateCatalogField(${r.id},'contesto',this.value)"></td>
      <td style="white-space:nowrap">
        <button class="btn small-btn btn-secondary" onclick="duplicaCatalogRow(${r.id})" style="padding:4px 8px;font-size:13px" title="Duplica">⎘</button>
        ${catalogRows.length > 1 ? `<button class="btn small-btn btn-danger" onclick="removeCatalogRow(${r.id})" style="padding:4px 8px;font-size:13px">✕</button>` : ''}
      </td>`);

    tbody.appendChild(tr);
  });
}

function updateCatalogField(id, field, val) {
  const r = catalogRows.find(r => r.id === id); if (r) r[field] = val;
}

function duplicaCatalogRow(id) {
  const src = catalogRows.find(r => r.id === id); if (!src) return;
  const newId = ++rowCounter;
  const idx   = catalogRows.findIndex(r => r.id === id);
  catalogRows.splice(idx + 1, 0, { ...src, id: newId, descrizione: '', filename: src.filename + '_dup' });
  renderCatalogRows();
}

function removeCatalogRow(id) {
  catalogRows = catalogRows.filter(r => r.id !== id);
  renderCatalogRows();
}

// ══════════════════════════════════════
// FASE 4 azioni — Salva / Aggiorna / Valida / Genera
// ══════════════════════════════════════

async function saveCatalog() {
  document.getElementById('btnSaveCatalog').disabled = true;
  const msg = document.getElementById('catalogSaveMsg');
  if (msg) { msg.textContent = '⏳ Salvataggio...'; msg.style.display = 'block'; }
  try {
    for (let i = 0; i < catalogRows.length; i++) {
      if (msg) msg.textContent = `Salvataggio voce ${i + 1}/${catalogRows.length}…`;
      const r = catalogRows[i];
      if (r.spItemId) {
        await updateCatalogoItem(r);
      } else {
        const created = await saveCatalogoItem(r, i);
        r.spItemId = created?.id || null;
      }
    }
    if (msg) { msg.textContent = '✅ Salvato'; setTimeout(() => { msg.style.display='none'; }, 3000); }
    // Reset flags dopo ogni salvataggio
    catalogoAggiornato = false;
    catalogoValidato   = false;
    updateButtonStates();
  } catch(e) {
    showMsg('❌ ' + e.message, 'error');
  } finally {
    document.getElementById('btnSaveCatalog').disabled = false;
  }
}

async function aggiornaCatalogo() {
  const btn = document.getElementById('btnAggiorna');
  if (btn) btn.disabled = true;
  showMsg('⏳ Aggiornamento in corso...', 'success');

  try {
    const drive      = await getDriveForFoto();
    const folderName = buildFolderName();

    // 1. Ordina per data cronologica
    catalogRows.sort((a, b) => {
      const da = a.data ? new Date(a.data + 'T00:00:00') : new Date(0);
      const db = b.data ? new Date(b.data + 'T00:00:00') : new Date(0);
      return da - db;
    });

    // 2. Calcola nuovi nomi per anno e identifica solo le rinominazioni necessarie
    const yearCounters = {};
    const renameTasks  = [];
    catalogRows.forEach(row => {
      const year = row.data ? new Date(row.data + 'T00:00:00').getFullYear() : new Date().getFullYear();
      if (!yearCounters[year]) yearCounters[year] = 0;
      const newName = buildNum(yearCounters[year], year) + '.jpg';
      yearCounters[year]++;
      if (row.filename && row.filename !== newName) {
        renameTasks.push({ row, oldName: row.filename, newName });
      }
      row.filename = newName;
    });

    // 3. Rinomina solo i file che hanno cambiato nome (PATCH minime)
    for (const task of renameTasks) {
      await renameFileOnDrive(drive.id, folderName, task.oldName, task.newName);
    }

    // 4. Aggiorna CatalogoFoto con nuovi nomi (solo le righe rinominate)
    for (const task of renameTasks) {
      if (task.row.spItemId) await updateCatalogoItem(task.row);
    }

    // 5. Salva anche le righe senza spItemId (nuove, non ancora mai salvate)
    const nuove = catalogRows.filter(r => !r.spItemId);
    for (const row of nuove) {
      await saveCatalogoItem(row);
    }

    renderCatalogRows();
    catalogoAggiornato = true;
    catalogoValidato   = false;
    updateButtonStates();

    const msg2 = renameTasks.length > 0
      ? `✅ ${renameTasks.length} file rinominati, ${catalogRows.length} righe riordinate.`
      : `✅ Ordine già corretto. ${catalogRows.length} righe nel catalogo.`;
    showMsg(msg2, 'success');

    // Ricarica thumbnail con i nuovi nomi file (quelli rinominati hanno URL non più validi)
    if (renameTasks.length > 0) {
      renameTasks.forEach(t => { t.row.thumbUrl = ''; }); // invalida vecchie URL
    }
    _loadThumbnailsFromSharePoint(catalogRows);

  } catch(e) {
    showMsg('❌ Aggiornamento fallito: ' + e.message, 'error');
    console.error(e);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function renameFileOnDrive(driveId, folderName, oldName, newName) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${encodeURIComponent(folderName + '/' + oldName)}:`,
    {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: newName })
    }
  );
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Rinomina ${oldName} → ${newName} fallita (${res.status}): ${txt.slice(0,100)}`);
  }
}

function validaCatalogo() {
  if (!catalogRows.length) {
    showMsg('⚠️ Nessuna foto nel catalogo da validare.', 'error');
    return;
  }
  catalogoValidato = true;
  updateButtonStates();
  showMsg('✅ Catalogo validato. Ora puoi generare il documento finale.', 'success');
}

function showCatalogoModelSelector() {
  if (!catalogoValidato) {
    showMsg('⚠️ Valida prima il catalogo.', 'error');
    return;
  }
  document.getElementById('modelSelectorModal').style.display = 'flex';
}
function closeCatalogoModal() {
  document.getElementById('modelSelectorModal').style.display = 'none';
}

async function esportaCatalogo(modello, formato) {
  closeCatalogoModal();
  const rows = catalogRows.map(r => ({
    num:         r.filename.replace(/\.jpg$/i, ''),
    operatore:   r.operatore   || '',
    dataFmt:     r.data ? new Date(r.data + 'T00:00:00').toLocaleDateString('it-IT') : '',
    localita:    r.localita    || '',
    descrizione: r.descrizione || '',
    vista:       r.vista       || '',
    contesto:    r.contesto    || '',
  }));

  if (modello === 'piemonte') {
    if (formato === 'pdf') openPrintWindow(buildPdfPiemonte(rows), 'elenco_foto_piemonte');
    else await buildWordPiemonte(rows);
  } else {
    if (formato === 'pdf') openPrintWindow(buildPdfVda(rows), 'elenco_foto_vda');
    else await buildWordVda(rows);
  }
}

// ══════════════════════════════════════
// BOTTOM SHEET — sostituzione singola foto
// ══════════════════════════════════════

function openSinglePicker(rowId)  { _pickerRowId = rowId; document.getElementById('fotoPickerOverlay').classList.add('active'); }
function closeFotoPicker() { document.getElementById('fotoPickerOverlay').classList.remove('active'); _pickerRowId = null; }
function triggerSinglePicker(source) {
  document.getElementById('fotoPickerOverlay').classList.remove('active');
  const rowId = _pickerRowId; _pickerRowId = null;
  if (!rowId) return;
  const input = document.getElementById(source === 'camera' ? 'singleCam' : 'singleGal');
  input.value = '';
  input.onchange = async () => {
    const file = input.files[0]; if (!file) return;
    const url  = await fileToDataUrl(file);
    const exif = await readExif(file);
    const r    = catalogRows.find(r => r.id === rowId);
    if (r) { r.file = file; r.previewUrl = url; r.exif = exif; if (exif.date) r.data = exif.date; }
    renderCatalogRows();
  };
  input.click();
}

// ══════════════════════════════════════
// GENERAZIONE CATALOGO — PIEMONTE + VDA
// ══════════════════════════════════════

function _catalogRowsToExport() {
  const src = catalogRows.length ? catalogRows : uploadedFotoData;
  return src.map(r => ({
    num:         (r.filename || r.num || '').replace(/\.jpg$/i, ''),
    operatore:   r.operatore   || '',
    dataFmt:     r.data ? new Date(r.data + 'T00:00:00').toLocaleDateString('it-IT') : (r.dataFmt || ''),
    localita:    r.localita    || '',
    descrizione: r.descrizione || '',
    vista:       r.vista       || '',
    contesto:    r.contesto    || '',
  }));
}

// Compatibilità bozza rapida
function generaPDF() { openPrintWindow(buildPdfPiemonte(_catalogRowsToExport()), 'elenco_bozza'); }

function buildPdfPiemonte(rows) {
  const comune = currentProj?.comune              || '';
  const prog   = currentProj?.codiceProgetto      || '';
  const comm   = currentProj?.committente         || '';
  const titolo = currentProj?.descrizioneProgetto || currentProj?.title || '';

  const tableRows = rows.map(r => `<tr>
    <td>${esc(r.num)}</td>
    <td style="white-space:pre-line">${esc(r.operatore)}\n${r.dataFmt}</td>
    <td style="text-align:left">${esc(r.localita)}</td>
    <td style="text-align:left">${esc(r.descrizione)}</td>
    <td>${r.vista || ''}</td>
    <td>${esc(r.contesto)}</td>
  </tr>`).join('');

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body{font-family:'Times New Roman',serif;font-size:11pt;margin:2cm}
  p{text-align:center;margin:3px 0}
  .hdr{font-size:12pt;font-weight:bold}
  .comune{font-size:13pt;font-weight:bold;margin-top:14px}
  .descr{font-size:11pt;font-weight:bold;font-style:italic;margin:4px 0}
  .subhdr,.elenco{font-size:11pt}
  .elenco{font-style:italic;margin-top:14px}
  table{width:100%;border-collapse:collapse;margin-top:16px;font-size:10pt}
  thead{display:table-header-group}
  th{background:#ccc;border:1px solid #666;padding:5px 7px;text-align:center;font-size:10pt}
  td{border:1px solid #888;padding:4px 7px;vertical-align:top;text-align:center}
  td:nth-child(3),td:nth-child(4){text-align:left}
  @media print{body{margin:1.5cm}@page{margin:1.5cm}}
</style></head><body>
<p class="hdr">SOPRINTENDENZA ARCHEOLOGICA DEL PIEMONTE</p>
<p class="hdr">ARCHIVIO FOTOGRAFICO BENI IMMOBILI</p>
${comune ? `<p class="comune">COMUNE DI ${esc(comune.toUpperCase())}</p>` : ''}
${(comm||prog) ? `<p class="subhdr">${[comm?esc(comm):'', prog?`PROG. ${esc(prog)}`:''].filter(Boolean).join(' ')}</p>` : ''}
${titolo ? `<p class="descr">${esc(titolo)}</p>` : ''}
<p class="elenco">ELENCO DELLE FOTOGRAFIE DIGITALI</p>
<table><thead><tr>
  <th style="width:14%">N. FOTO</th>
  <th style="width:16%">OPERATORE<br>E DATA</th>
  <th style="width:18%">LOCALITÀ</th>
  <th style="width:30%">DESCRIZIONE</th>
  <th style="width:8%">VISTA<br>DA</th>
  <th style="width:14%">CONTESTO</th>
</tr></thead><tbody>${tableRows}</tbody></table>
</body></html>`;
}

function buildPdfVda(rows) {
  const comune = currentProj?.comune              || '';
  const codice = currentProj?.codiceSito          || '';
  const denom  = currentProj?.descrizioneProgetto || currentProj?.title || '';

  const tableRows = rows.map(r => {
    const descParts = [r.localita, r.contesto, r.descrizione, r.vista ? `da ${r.vista}` : ''].filter(Boolean);
    return `<tr>
      <td>${r.dataFmt}</td>
      <td>${esc(r.operatore)}</td>
      <td>${esc(r.num)}</td>
      <td style="text-align:left">${esc(descParts.join(', '))}</td>
    </tr>`;
  }).join('');

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body{font-family:'Times New Roman',serif;font-size:11pt;margin:1.5cm}
  p{margin:4px 0}
  .titolo{font-size:13pt;font-weight:bold;text-align:center;margin-bottom:12px}
  table{width:100%;border-collapse:collapse;margin-top:16px;font-size:10pt}
  thead{display:table-header-group}
  th{background:#ccc;border:1px solid #666;padding:5px 7px;text-align:center;font-size:10pt}
  td{border:1px solid #888;padding:4px 7px;vertical-align:top;text-align:center}
  td:nth-child(4){text-align:left}
  @media print{body{margin:1cm}@page{size:A4 landscape;margin:1cm}}
</style></head><body>
<p class="titolo">ELENCO DOCUMENTAZIONE FOTOGRAFICA</p>
<p><strong>COMUNE:</strong> ${esc(comune)}</p>
<p><strong>LOCALITÀ o SPAZIO VIABILISTICO:</strong></p>
<p><strong>CODICE:</strong> ${esc(codice)}</p>
<p><strong>DENOMINAZIONE:</strong> ${esc(denom)}</p>
<table><thead><tr>
  <th style="width:11%">DATA</th>
  <th style="width:15%">AUTORE</th>
  <th style="width:20%">NOME FILE</th>
  <th style="width:54%">DESCRIZIONE</th>
</tr></thead><tbody>${tableRows}</tbody></table>
</body></html>`;
}

async function buildWordPiemonte(rows) {
  try {
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, BorderStyle } = docx;
    const comune = currentProj?.comune||'', prog = currentProj?.codiceProgetto||'';
    const comm   = currentProj?.committente||'';
    const titolo = currentProj?.descrizioneProgetto || currentProj?.title || '';
    const bdr = { style:BorderStyle.SINGLE, size:6, color:'888888' };
    const cb  = { top:bdr, bottom:bdr, left:bdr, right:bdr };
    const hC  = (t,w) => new TableCell({ width:{size:w,type:WidthType.DXA}, borders:cb, shading:{fill:'CCCCCC'},
      children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:t,bold:true,size:18})]})] });
    const dC  = (t,al=AlignmentType.CENTER,w=1000) => new TableCell({ width:{size:w,type:WidthType.DXA}, borders:cb,
      children:[new Paragraph({alignment:al, children:[new TextRun({text:t||'',size:18})]})] });
    const tRows = [
      new TableRow({children:[hC('N. FOTO',1700),hC('OPERATORE\nE DATA',2100),hC('LOCALITÀ',2100),hC('DESCRIZIONE',3600),hC('VISTA DA',900),hC('CONTESTO',1900)]}),
      ...rows.map(r => new TableRow({children:[
        dC(r.num,AlignmentType.CENTER,1700), dC(`${r.operatore}\n${r.dataFmt}`,AlignmentType.CENTER,2100),
        dC(r.localita,AlignmentType.LEFT,2100), dC(r.descrizione,AlignmentType.LEFT,3600),
        dC(r.vista,AlignmentType.CENTER,900), dC(r.contesto,AlignmentType.CENTER,1900),
      ]}))
    ];
    const cp = (t,bold=false,sz=22) => new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:t,bold,size:sz})]});
    const doc = new Document({ sections:[{ children:[
      cp('SOPRINTENDENZA ARCHEOLOGICA DEL PIEMONTE',true,22), cp('ARCHIVIO FOTOGRAFICO BENI IMMOBILI',true,22), cp(''),
      ...(comune?[cp(`COMUNE DI ${comune.toUpperCase()}`,true,26)]:[]), cp(''),
      ...((comm||prog)?[cp(`${comm}${prog?` PROG. ${prog}`:''}`,false,22)]:[]),
      ...(titolo?[cp(titolo,true,22)]:[]), cp(''), cp('ELENCO DELLE FOTOGRAFIE DIGITALI',false,22), cp(''),
      new Table({width:{size:12300,type:WidthType.DXA}, rows:tRows}),
    ]}]});
    _downloadBlob(await Packer.toBlob(doc), `elenco_piemonte_${sanitize(currentProj?.title||'')}.docx`);
  } catch(e) { alert('Errore Word Piemonte: '+e.message); console.error(e); }
}

async function buildWordVda(rows) {
  try {
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, BorderStyle, PageOrientation } = docx;
    const comune = currentProj?.comune||'', codice = currentProj?.codiceSito||'';
    const denom  = currentProj?.descrizioneProgetto || currentProj?.title || '';
    const bdr = { style:BorderStyle.SINGLE, size:6, color:'888888' };
    const cb  = { top:bdr, bottom:bdr, left:bdr, right:bdr };
    const hC  = (t,w) => new TableCell({ width:{size:w,type:WidthType.DXA}, borders:cb, shading:{fill:'CCCCCC'},
      children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:t,bold:true,size:18})]})] });
    const dC  = (t,al=AlignmentType.CENTER,w=1000) => new TableCell({ width:{size:w,type:WidthType.DXA}, borders:cb,
      children:[new Paragraph({alignment:al, children:[new TextRun({text:t||'',size:18})]})] });
    const tRows = [
      new TableRow({children:[hC('DATA',1400),hC('AUTORE',2000),hC('NOME FILE',2600),hC('DESCRIZIONE',8300)]}),
      ...rows.map(r => {
        const dp = [r.localita,r.contesto,r.descrizione,r.vista?`da ${r.vista}`:''].filter(Boolean);
        return new TableRow({children:[dC(r.dataFmt,AlignmentType.CENTER,1400),dC(r.operatore,AlignmentType.CENTER,2000),dC(r.num,AlignmentType.CENTER,2600),dC(dp.join(', '),AlignmentType.LEFT,8300)]});
      })
    ];
    const lbl = (label,val) => new Paragraph({children:[new TextRun({text:label+' ',bold:true,size:22}),new TextRun({text:val||'',size:22})]});
    const doc = new Document({ sections:[{ properties:{page:{size:{orientation:PageOrientation.LANDSCAPE}}}, children:[
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:'ELENCO DOCUMENTAZIONE FOTOGRAFICA',bold:true,size:26})]}),
      new Paragraph({text:''}),
      lbl('COMUNE:',comune), lbl('LOCALITÀ o SPAZIO VIABILISTICO:',''), lbl('CODICE:',codice), lbl('DENOMINAZIONE:',denom),
      new Paragraph({text:''}),
      new Table({width:{size:14300,type:WidthType.DXA}, rows:tRows}),
    ]}]});
    _downloadBlob(await Packer.toBlob(doc), `elenco_vda_${sanitize(currentProj?.title||'')}.docx`);
  } catch(e) { alert('Errore Word VdA: '+e.message); console.error(e); }
}

function _downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = Object.assign(document.createElement('a'), {href:url, download:filename});
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ══════════════════════════════════════
// MODALITÀ RIVEDI CATALOGO
// ══════════════════════════════════════

async function enterReviewMode() {
  if (!currentProj) return;
  syncInfoBars();
  showMsg('⏳ Caricamento catalogo da SharePoint...', 'success');
  try {
    const items = await getExistingCatalogItems();
    if (!items.length) {
      showMsg('ℹ️ Nessuna foto ancora catalogata per questo cantiere.', 'success');
      return;
    }

    // Sincronizzazione silente: rimuovi righe orfane (file eliminati da SP)
    const cleaned = await _sincronizzaSilente(items);

    catalogRows = cleaned;
    catalogoAggiornato = false;
    catalogoValidato   = false;
    syncInfoBars();
    renderCatalogRows();
    updateButtonStates();
    goToPhase('phase-catalog', 4);

    const rimossi = items.length - cleaned.length;
    if (rimossi > 0) {
      showMsg(`📋 Catalogo caricato: ${cleaned.length} foto (${rimossi} rimosse perché non più presenti su SharePoint).`, 'success');
    } else {
      showMsg(`📋 Catalogo caricato: ${cleaned.length} foto.`, 'success');
    }

    // Carica thumbnail da SharePoint in background
    _loadThumbnailsFromSharePoint(cleaned);

  } catch(e) {
    showMsg('❌ Errore caricamento: ' + e.message, 'error');
  }
}

// Sincronizzazione leggera: legge i file presenti nella cartella SP
// ed elimina da CatalogoFoto le righe orfane. Restituisce le righe valide.
async function _sincronizzaSilente(items) {
  try {
    const drive      = await getDriveForFoto();
    const folderName = buildFolderName();
    const files      = await getDriveChildren(drive.id, folderName, 'name');
    const filesOnDisk = new Set(files
      .map(f => f.name.toLowerCase())
      .filter(n => /\.(jpg|jpeg|png|heic)$/.test(n)));

    const orphans = items.filter(r => {
      const fname = ((r.filename || '') + (r.filename && r.filename.includes('.') ? '' : '.jpg')).toLowerCase();
      return fname && !filesOnDisk.has(fname);
    });

    // Cancella le righe orfane da CatalogoFoto
    for (const row of orphans) {
      if (row.spItemId) {
        await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${catalogoListId}/items/${row.spItemId}`,
          { method: 'DELETE', headers: { 'Authorization': `Bearer ${accessToken}` } }
        ).catch(() => {}); // ignora errori singoli
      }
    }

    return items.filter(r => !orphans.includes(r));
  } catch(e) {
    console.warn('Sincronizzazione silente fallita:', e.message);
    return items; // in caso di errore ritorna tutto
  }
}

async function _loadThumbnailsFromSharePoint(items) {
  try {
    const drive      = await getDriveForFoto();
    const folderName = buildFolderName();
    for (const item of items) {
      try {
        const fname = item.filename ? (item.filename.endsWith('.jpg') ? item.filename : item.filename + '.jpg') : '';
        if (!fname) continue;
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${drive.id}/root:/${encodeURIComponent(folderName + '/' + fname)}:/thumbnails/0/medium`,
          { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );
        if (!res.ok) continue;
        const td = await res.json();
        if (td.url) {
          // Salva su row in modo persistente (sopravvive ai re-render)
          item.thumbUrl = td.url;
          // Aggiorna anche l'img nel DOM se presente
          const cell = document.querySelector(`[data-filename="${CSS.escape(item.filename)}"]`);
          if (cell) {
            const img = cell.querySelector('img');
            if (img) img.src = td.url;
            const ph = cell.querySelector('div');
            if (ph && !img) {
              const newImg = document.createElement('img');
              newImg.src = td.url;
              newImg.style.cssText = 'width:56px;height:56px;object-fit:cover;border-radius:6px;border:1px solid #ddd;cursor:pointer';
              newImg.addEventListener('click', () => openLightboxFull(item));
              cell.insertBefore(newImg, ph);
              ph.remove();
            }
          }
        }
      } catch(e) { /* skip */ }
    }
  } catch(e) {
    console.warn('Thumbnail SP non disponibili:', e.message);
  }
}

// ══════════════════════════════════════
// HELPERS
// ══════════════════════════════════════

function buildFolderName() {
  const sito   = sanitize(currentProj?.codiceSito   || '');
  const comm   = sanitize(currentProj?.committente  || '');
  const yy     = String(new Date().getFullYear()).slice(2);
  const comune = sanitize(currentProj?.comune       || '');
  const prog   = sanitize(currentProj?.codiceProgetto || '');
  return `${sito}_${comm}${yy}_${comune}_${prog}_Documentazione fotografica`;
}

function buildNum(yearIndex, year) {
  const base = buildFolderName().replace(/_Documentazione fotografica$/i, '');
  return `${base}_${String(yearIndex + 1).padStart(3,'0')}`;
}

function buildNumPreview(i, year) {
  const base = buildFolderName().replace(/_Documentazione fotografica$/i, '');
  return `${base}_${String(i + 1).padStart(3,'0')}`;
}

async function getExistingPhotoCountForYear(driveId, folderName, year) {
  try {
    const base   = buildFolderName().replace(/_Documentazione fotografica$/i, '');
    const prefix = `${base}_`.toLowerCase();
    const files  = await getDriveChildren(driveId, folderName, 'name');
    return files.filter(f =>
      f.name.toLowerCase().startsWith(prefix) &&
      /\.(jpg|jpeg|png|heic)$/i.test(f.name)
    ).length;
  } catch(e) { return 0; }
}

function sanitize(s) {
  return (s || '').trim().replace(/\s+/g, '').replace(/[^a-zA-Z0-9_-]/g, '');
}

function sortByExifDate(a, b) {
  return (a.exif?.dateObj || new Date(0)) - (b.exif?.dateObj || new Date(0));
}

async function readExif(file) {
  try {
    const data = await exifr.parse(file, {
      pick: ['DateTimeOriginal','GPSLatitude','GPSLongitude','GPSLatitudeRef','GPSLongitudeRef']
    });
    if (!data) return {};
    const result = {};
    if (data.DateTimeOriginal) {
      const d = new Date(data.DateTimeOriginal);
      if (!isNaN(d)) {
        result.dateObj = d;
        result.date    = d.toISOString().split('T')[0];
        result.dateStr = d.toLocaleDateString('it-IT');
      }
    }
    if (data.GPSLatitude && data.GPSLongitude) {
      let lat = data.GPSLatitude[0] + data.GPSLatitude[1]/60 + data.GPSLatitude[2]/3600;
      let lng = data.GPSLongitude[0] + data.GPSLongitude[1]/60 + data.GPSLongitude[2]/3600;
      if (data.GPSLatitudeRef  === 'S') lat = -lat;
      if (data.GPSLongitudeRef === 'W') lng = -lng;
      result.lat = lat; result.lng = lng;
    }
    return result;
  } catch(e) { return {}; }
}

async function reverseGeocode(lat, lng) {
  try {
    const r = await fetch(
      `https://nominatim.openstreetmap.org/reverse?lat=${lat}&lon=${lng}&format=json&accept-language=it`,
      { headers: { 'User-Agent': 'Akhet-PostScavo/1.0' } }
    );
    const d = await r.json();
    const a = d.address || {};
    const parts = [a.road || a.pedestrian || a.path || a.footway, a.hamlet || a.suburb || a.village || a.town || a.city].filter(Boolean);
    return parts.join(', ') || d.display_name?.split(',').slice(0,2).join(', ') || '';
  } catch(e) { return ''; }
}

function compImg(file, maxPx = 1200, q = 0.8) {
  return new Promise((ok, err) => {
    const r = new FileReader();
    r.onload = e => {
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas');
        let w = img.width, h = img.height;
        if (w > h && w > maxPx) { h *= maxPx/w; w = maxPx; } else if (h > maxPx) { w *= maxPx/h; h = maxPx; }
        c.width = w; c.height = h;
        c.getContext('2d').drawImage(img, 0, 0, w, h);
        c.toBlob(ok, 'image/jpeg', q);
      };
      img.onerror = err; img.src = e.target.result;
    };
    r.onerror = err; r.readAsDataURL(file);
  });
}

function fileToDataUrl(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => {
      const result = e.target.result;
      console.log(`fileToDataUrl ${file.name}: tipo=${typeof result}, lunghezza=${result?.length || 0}, inizio=${result?.slice(0,30) || 'VUOTO'}`);
      res(result || '');
    };
    r.onerror = e => { console.error('FileReader errore:', e); res(''); };
    r.readAsDataURL(file);
  });
}

function esc(s) { return (s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

// Apre HTML in una nuova finestra per la stampa.
// Se il popup è bloccato dal browser, scarica il file come blob.
function openPrintWindow(html, fallbackFileName) {
  const w = window.open('', '_blank');
  if (w && w.document) {
    w.document.write(html);
    w.document.close();
    setTimeout(() => { try { w.print(); } catch(e) {} }, 700);
  } else {
    // Popup bloccato → scarica come file HTML apribile nel browser
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = (fallbackFileName || 'elenco') + '.html';
    a.target   = '_blank';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    showMsg('ℹ️ Il popup è stato bloccato dal browser — il file è stato scaricato come HTML. Aprilo e usa Stampa (Ctrl+P) per salvarlo in PDF.', 'success');
  }
}

function showMsg(msg, type) {
  const d = document.getElementById('appMsg');
  d.innerHTML = `<div class="${type==='error'?'msg-error':'msg-success'}" style="margin-bottom:15px">${msg}</div>`;
  setTimeout(() => d.innerHTML = '', 8000);
}

function openLightbox(src) {
  document.getElementById('lightboxImg').src = src;
  document.getElementById('lightbox').classList.add('active');
}
function closeLightbox() { document.getElementById('lightbox').classList.remove('active'); }

async function openLightboxFull(row) {
  const lb    = document.getElementById('lightbox');
  const img   = document.getElementById('lightboxImg');

  // Se c'è un previewUrl locale (upload corrente) usalo direttamente
  if (row.previewUrl && row.previewUrl.startsWith('data:')) {
    img.src = row.previewUrl;
    lb.classList.add('active');
    return;
  }

  // Altrimenti carica la versione grande da SharePoint
  lb.classList.add('active');
  img.src = row.previewUrl || ''; // mostra thumbnail mentre carica

  try {
    const drive      = await getDriveForFoto();
    const folderName = buildFolderName();
    const fname      = row.filename.endsWith('.jpg') ? row.filename : row.filename + '.jpg';

    // Prima prova thumbnail "large" (veloce)
    const trRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${drive.id}/root:/${encodeURIComponent(folderName + '/' + fname)}:/thumbnails/0/large`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    if (trRes.ok) {
      const td = await trRes.json();
      if (td.url) { img.src = td.url; return; }
    }

    // Fallback: scarica il file direttamente
    const dlRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${drive.id}/root:/${encodeURIComponent(folderName + '/' + fname)}:/content`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    if (dlRes.ok) {
      const blob = await dlRes.blob();
      img.src = URL.createObjectURL(blob);
    }
  } catch(e) {
    console.warn('openLightboxFull errore:', e.message);
  }
}

function resetAll() {
  // Reset stato
  currentProj = null;
  bulkFiles = [];
  catalogRows = [];
  rowCounter = 0;
  uploadedFotoData = [];
  catalogoAggiornato = false;
  catalogoValidato   = false;

  // Reset UI cantiere
  const sel = document.getElementById('selCantiere');
  if (sel) sel.value = '';
  document.querySelectorAll('.cantiere-card').forEach(c => c.classList.remove('selected'));
  ['cantInfoBar','cantInfoBar2','cantInfoBar4','uploadLegend'].forEach(id => {
    const el = document.getElementById(id); if(el) el.style.display = 'none';
  });
  document.getElementById('modeSelector').style.display = 'none';

  // Reset UI upload
  const bsu = document.getElementById('btnStartUpload'); if(bsu) bsu.disabled = false;
  const bsc = document.getElementById('btnSaveCatalog'); if(bsc) bsc.disabled = false;
  document.getElementById('uploadActions').style.display = 'none';
  document.getElementById('previewGrid').innerHTML = '';
  document.getElementById('progressBar').style.width = '0%';
  document.getElementById('uploadLog').innerHTML = '';

  // Reset UI catalogo
  const tbody = document.getElementById('catalogRows'); if(tbody) tbody.innerHTML = '';
  const csm   = document.getElementById('catalogSaveMsg'); if(csm) csm.style.display = 'none';

  // Aggiorna bottoni ruolo
  updateButtonStates();

  goToPhase('phase-cantiere', 1);
  window.scrollTo(0, 0);
}

// ══════════════════════════════════════
// AVVIO
// ══════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  // initializeApp() viene chiamato dal flusso MSAL dopo il login
  tavInit();

  // Drag & drop foto
  const area = document.getElementById('uploadArea');
  if (area) {
    area.addEventListener('dragover',  e => { e.preventDefault(); area.classList.add('drag-over'); });
    area.addEventListener('dragleave', ()  => area.classList.remove('drag-over'));
    area.addEventListener('drop', e => {
      e.preventDefault(); area.classList.remove('drag-over');
      const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
      if (!files.length) return;
      const dt = new DataTransfer(); files.forEach(f => dt.items.add(f));
      const inp = document.getElementById('bulkInput'); inp.files = dt.files;
      onBulkSelect({ target: inp });
    });
  }

  // Drag & drop tavole PDF
  const tavArea = document.getElementById('tavUploadArea');
  if (tavArea) {
    tavArea.addEventListener('dragover',  e => { e.preventDefault(); tavArea.classList.add('drag-over'); });
    tavArea.addEventListener('dragleave', ()  => tavArea.classList.remove('drag-over'));
    tavArea.addEventListener('drop', e => {
      e.preventDefault(); tavArea.classList.remove('drag-over');
      const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf');
      if (!files.length) return;
      const dt = new DataTransfer(); files.forEach(f => dt.items.add(f));
      const inp = document.getElementById('tavInput'); inp.files = dt.files;
      tavOnSelect({ target: inp });
    });
  }
});

// ══════════════════════════════════════════════════════
// SEZIONE ELENCO TAVOLE
// ══════════════════════════════════════════════════════

// Stato tavole
let tavProj      = null;
let tavFiles     = [];   // { id, file, fileName, text, data, autore, descrizione, pdfDoc }
let tavCounter   = 0;

// Inizializza PDF.js worker
function tavInit() {
  // Il dropdown tavSelCantiere viene popolato da loadCantieri() dopo il login SharePoint
}

// ── Navigazione ────────────────────────────────────────

function tavGoToPhase(phaseId, stepNum) {
  document.querySelectorAll('.tav-phase').forEach(p => p.classList.remove('active'));
  document.getElementById(phaseId).classList.add('active');
  document.querySelectorAll('#tavStepIndicator .step').forEach((s, i) => {
    s.classList.remove('active','done');
    if (i + 1 < stepNum) s.classList.add('done');
    else if (i + 1 === stepNum) s.classList.add('active');
  });
  window.scrollTo(0, 0);
}

function tavSyncInfoBars() {
  const html = document.getElementById('tavCantInfoBar').innerHTML;
  ['tavCantInfoBar2','tavCantInfoBar3'].forEach(id => {
    const el = document.getElementById(id); if (el) el.innerHTML = html;
  });
}

// ── Fase 1: selezione cantiere ─────────────────────────

function tavOnCantiereChange() {
  const id = document.getElementById('tavSelCantiere').value;
  tavProj = cantieriData.find(c => c.id === id) || null;
  const bar = document.getElementById('tavCantInfoBar');
  const btn = document.getElementById('tavBtnToUpload');
  if (!tavProj) { bar.style.display='none'; btn.style.display='none'; return; }
  bar.style.display = 'block';
  bar.innerHTML = `<strong>${tavProj.title}</strong>`
    + (tavProj.committente ? ` &nbsp;|&nbsp; ${tavProj.committente}` : '')
    + (tavProj.codiceSito  ? ` &nbsp;|&nbsp; Cod. sito: <strong>${tavProj.codiceSito}</strong>` : '')
    + (tavProj.codiceProgetto ? ` &nbsp;|&nbsp; Prog. <strong>${tavProj.codiceProgetto}</strong>` : '');
  btn.style.display = 'inline-block';
}

// ── Fase 2: upload PDF ─────────────────────────────────

async function tavOnSelect(event) {
  const files = Array.from(event.target.files);
  if (!files.length) return;
  event.target.value = '';
  const grid = document.getElementById('tavPreviewGrid');

  for (const file of files) {
    const id = ++tavCounter;
    tavFiles.push({ id, file, fileName: file.name, text: '', data: '', autore: '', descrizione: '' });
    grid.appendChild(buildTavCard(id, file.name));
  }

  document.getElementById('tavUploadActions').style.display = 'flex';
  document.getElementById('tavCount').textContent = `${tavFiles.length} tavol${tavFiles.length === 1 ? 'a' : 'e'} selezionat${tavFiles.length === 1 ? 'a' : 'e'}`;
}

function buildTavCard(id, fileName) {
  const card = document.createElement('div');
  card.className = 'tav-card processing';
  card.id = 'tcard-' + id;
  card.innerHTML = `
    <div class="tav-badge processing" id="tbadge-${id}">⏳</div>
    <div class="tav-thumb-canvas" id="tcanvas-wrap-${id}" style="display:flex;align-items:center;justify-content:center;background:#f0ece4">
      <span style="font-size:36px">📄</span>
    </div>
    <button class="tav-remove" onclick="tavRemove(${id})">✕</button>
    <div class="tav-card-meta">
      <div class="tav-card-name">${esc(fileName)}</div>
      <div id="tmeta-${id}" style="color:#aaa;font-size:10px">In attesa...</div>
    </div>`;
  return card;
}

function tavRemove(id) {
  tavFiles = tavFiles.filter(f => f.id !== id);
  document.getElementById('tcard-' + id)?.remove();
  const cnt = document.getElementById('tavCount');
  cnt.textContent = tavFiles.length
    ? `${tavFiles.length} tavol${tavFiles.length===1?'a':'e'} selezionat${tavFiles.length===1?'a':'e'}`
    : '';
  if (!tavFiles.length) document.getElementById('tavUploadActions').style.display = 'none';
}

// ── Fase 3: lettura PDF + estrazione testo ──────────────

async function tavGoToPreview() {
  if (!tavFiles.length) { alert('Carica almeno un PDF.'); return; }
  tavSyncInfoBars();
  document.getElementById('btnTavNext').disabled = true;
  tavGoToPhase('tav-phase-reading', 3);

  const bar     = document.getElementById('tavProgressBar');
  const msg     = document.getElementById('tavReadMsg');
  const counter = document.getElementById('tavReadCounter');

  for (let i = 0; i < tavFiles.length; i++) {
    const item = tavFiles[i];
    bar.style.width = Math.round((i / tavFiles.length) * 100) + '%';
    counter.textContent = `Lettura ${i+1} di ${tavFiles.length}: ${item.fileName}`;
    msg.textContent = 'Apertura PDF...';

    try {
      const buf = await item.file.arrayBuffer();

      // Opzioni per performance con file pesanti
      const loadTask = pdfjsLib.getDocument({
        data:               buf,
        disableFontFace:    true,   // non scarica font esterni
        isEvalSupported:    false,  // sicurezza
        disableRange:       false,
        disableStream:      false,
      });
      const pdf = await loadTask.promise;
      item.pdfDoc = pdf;

      // Estrai solo la data dal cartiglio
      msg.textContent = `Lettura data ${i+1}/${tavFiles.length}...`;
      item.data = await estraiSoloData(pdf);
      // autore e descrizione restano vuoti → compilazione libera dall'utente

      // Miniatura compressa (scala bassa per file pesanti)
      msg.textContent = `Anteprima ${i+1}/${tavFiles.length}...`;
      const page1 = await pdf.getPage(1);
      const vp    = page1.getViewport({ scale: 0.22 });
      const canvas = document.createElement('canvas');
      canvas.width  = vp.width;
      canvas.height = vp.height;
      canvas.style.width  = '100%';
      canvas.style.cursor = 'pointer';
      canvas.onclick = () => openTavPreview(item.id);
      await page1.render({
        canvasContext: canvas.getContext('2d'),
        viewport:      vp,
        intent:        'print',   // più veloce di 'display'
      }).promise;

      item.thumbDataUrl = canvas.toDataURL('image/jpeg', 0.55);

      const wrap = document.getElementById('tcanvas-wrap-' + item.id);
      if (wrap) { wrap.innerHTML = ''; wrap.appendChild(canvas); }

      // Aggiorna card
      const badge  = document.getElementById('tbadge-' + item.id);
      const metaEl = document.getElementById('tmeta-' + item.id);
      const card   = document.getElementById('tcard-' + item.id);
      if (badge)  { badge.className = 'tav-badge done-badge'; badge.textContent = '✓'; }
      if (card)   card.className = 'tav-card done-card';
      if (metaEl) metaEl.textContent = item.data || '—';

    } catch(e) {
      console.error('Errore lettura PDF:', item.fileName, e);
      const metaEl = document.getElementById('tmeta-' + item.id);
      if (metaEl) metaEl.textContent = '⚠️ Errore lettura';
      item.data = '';
    }
  }

  bar.style.width = '100%';
  renderTavRows();
  tavGoToPhase('tav-phase-preview', 3);
  document.getElementById('btnTavNext').disabled = false;
}

// ── Estrazione SOLO DATA dal cartiglio ─────────────────────

async function estraiSoloData(pdfDoc) {
  try {
    const page   = await pdfDoc.getPage(1);
    const content = await page.getTextContent();

    const items = content.items
      .filter(it => it.str && it.str.trim().length > 0)
      .map(it => ({
        str: it.str.trim(),
        x:   it.transform[4],
        y:   it.transform[5],
        w:   it.width || 0,
      }));

    const rows = groupRows(items, 6);

    // Strategia 1: trova header "Data" e prendi il valore nella riga sotto
    const dataHdrIdx = rows.findIndex(row => row.some(it => /^data$/i.test(it.str)));
    if (dataHdrIdx >= 0) {
      const dataHdr = rows[dataHdrIdx].find(it => /^data$/i.test(it.str));
      const valRow  = rows[dataHdrIdx + 1];
      if (dataHdr && valRow) {
        const closest = closestByX(valRow, dataHdr.x);
        if (closest && /\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{4}/.test(closest.str)) {
          return closest.str.replace(/[\.\-]/g, '/');
        }
      }
    }

    // Strategia 2: prima data trovata nel testo (qualsiasi formato gg/mm/aaaa)
    const allText = items.map(it => it.str).join(' ');
    const m = allText.match(/\b(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{4})\b/);
    if (m) return `${m[1].padStart(2,'0')}/${m[2].padStart(2,'0')}/${m[3]}`;

    // Strategia 3: metadati PDF
    const meta = await pdfDoc.getMetadata().catch(() => ({}));
    const raw  = meta?.info?.CreationDate || '';
    const mm   = raw.match(/D:(\d{4})(\d{2})(\d{2})/);
    if (mm) return `${mm[3]}/${mm[2]}/${mm[1]}`;

  } catch(e) {
    console.warn('estraiSoloData error:', e);
  }
  return '';
}

// ── Helpers posizionali ─────────────────────────────────────

function groupRows(items, tolerance) {
  if (!items.length) return [];
  const sorted = [...items].sort((a,b) => b.y - a.y || a.x - b.x);
  const rows = [[sorted[0]]];
  for (let i = 1; i < sorted.length; i++) {
    const lastRow = rows[rows.length - 1];
    const refY    = lastRow[0].y;
    if (Math.abs(sorted[i].y - refY) <= tolerance) {
      lastRow.push(sorted[i]);
    } else {
      rows.push([sorted[i]]);
    }
  }
  return rows.map(r => r.sort((a,b) => a.x - b.x));
}

function closestByX(rowItems, targetX) {
  if (!rowItems.length) return null;
  return rowItems.reduce((best, it) =>
    Math.abs(it.x - targetX) < Math.abs(best.x - targetX) ? it : best
  , rowItems[0]);
}

// ── Fase 4: tabella preview ─────────────────────────────

function renderTavRows() {
  const tbody = document.getElementById('tavRows');
  tbody.innerHTML = '';
  tavFiles.forEach((r, i) => {
    const tr = document.createElement('tr');
    tr.id = 'tavrow-' + r.id;
    tr.innerHTML = `
      <td style="text-align:center">
        <button onclick="openTavPreview(${r.id})" style="background:none;border:none;font-size:18px;cursor:pointer" title="Anteprima">👁️</button>
      </td>
      <td><input type="text" class="sm" value="${esc(r.data)}" placeholder="gg/mm/aaaa" oninput="tavUpdateField(${r.id},'data',this.value)" style="min-width:90px"></td>
      <td><input type="text" class="sm" value="${esc(r.autore)}" placeholder="Autore" oninput="tavUpdateField(${r.id},'autore',this.value)" style="min-width:110px"></td>
      <td class="num-cell" style="font-size:11px;word-break:break-all">${esc(r.fileName)}</td>
      <td><input type="text" class="sm" value="${esc(r.descrizione)}" placeholder="Descrizione / Titolo tavola" oninput="tavUpdateField(${r.id},'descrizione',this.value)" style="min-width:200px"></td>
      <td style="white-space:nowrap">
        ${tavFiles.length > 1 ? `<button class="btn small-btn btn-danger" onclick="tavRemoveRow(${r.id})" style="padding:3px 7px;font-size:12px">✕</button>` : ''}
      </td>`;
    tbody.appendChild(tr);
  });
}

function tavUpdateField(id, field, val) {
  const r = tavFiles.find(f => f.id === id); if (r) r[field] = val;
}

function tavRemoveRow(id) {
  tavFiles = tavFiles.filter(f => f.id !== id);
  document.getElementById('tavrow-' + id)?.remove();
}

// ── Preview modale ──────────────────────────────────────

async function openTavPreview(id) {
  const item = tavFiles.find(f => f.id === id);
  if (!item) return;
  const modal  = document.getElementById('tavPreviewModal');
  const canvas = document.getElementById('tavPreviewCanvas');
  const nameEl = document.getElementById('tavPreviewName');
  nameEl.textContent = item.fileName;
  canvas.width = 0; canvas.height = 0; // reset
  modal.classList.add('active');

  try {
    const pdf  = item.pdfDoc || await pdfjsLib.getDocument({ data: await item.file.arrayBuffer() }).promise;
    const page = await pdf.getPage(1);
    const vp   = page.getViewport({ scale: 1.5 });
    canvas.width  = vp.width;
    canvas.height = vp.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
  } catch(e) {
    canvas.width = 300; canvas.height = 100;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fee'; ctx.fillRect(0,0,300,100);
    ctx.fillStyle = '#c00'; ctx.font = '14px Arial';
    ctx.fillText('Errore rendering PDF', 20, 55);
  }
}

function closeTavPreview() {
  document.getElementById('tavPreviewModal').classList.remove('active');
}

// ── Genera e scarica PDF elenco ─────────────────────────

function scaricaElencoTavole() {
  if (!tavFiles.length) { alert('Nessuna tavola da esportare.'); return; }

  const comune  = tavProj?.comune         || '';
  const prog    = tavProj?.codiceProgetto  || '';
  const comm    = tavProj?.committente     || '';
  const titolo  = tavProj?.title           || '';

  // Costruisci righe tabella con virgolette per valori ripetuti
  let prevData = null, prevAutore = null;
  const tableRows = tavFiles.map(r => {
    const newData   = r.data   !== prevData;
    const newAutore = r.autore !== prevAutore;
    const dataCell   = newData   ? esc(r.data)   : '"';
    const autoreCell = newAutore ? esc(r.autore)  : '"';
    prevData = r.data; prevAutore = r.autore;
    return `<tr>
      <td>${dataCell}</td>
      <td>${autoreCell}</td>
      <td style="text-align:left;word-break:break-word">${esc(r.fileName)}</td>
      <td style="text-align:left">${esc(r.descrizione)}</td>
    </tr>`;
  }).join('');

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body { font-family: 'Times New Roman', Times, serif; font-size: 11pt; margin: 2cm; }
  p.center { text-align: center; margin: 3px 0; }
  p.title-lg { font-size: 12pt; font-weight: bold; text-align: center; margin: 3px 0; }
  p.title-sm { font-size: 11pt; text-align: center; margin: 3px 0; }
  p.italic   { font-style: italic; text-align: center; margin: 3px 0; }
  table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 10pt; }
  thead { display: table-header-group; }
  th { background: #cccccc; border: 1px solid #666; padding: 5px 7px; text-align: center; font-size: 10pt; }
  td { border: 1px solid #888; padding: 4px 7px; vertical-align: top; text-align: center; }
  small { font-size: 9pt; }
  @media print { body { margin: 1.5cm; } @page { size: A4 portrait; margin: 1.5cm; } }
</style>
</head><body>
<p class="title-lg">SOPRINTENDENZA ARCHEOLOGICA DEL PIEMONTE</p>
<p class="title-lg">ARCHIVIO FOTOGRAFICO BENI IMMOBILI</p>
<br>
${comune ? `<p class="title-sm">COMUNE DI ${esc(comune.toUpperCase())}</p><br>` : ''}
${prog   ? `<p class="title-sm">${esc(comm)} PROG. ${esc(prog)}</p>` : ''}
${titolo ? `<p class="title-sm">${esc(titolo)}</p>` : ''}
<br>
<p class="italic">ELENCO DOCUMENTAZIONE GRAFICA</p>
<table>
  <thead>
    <tr>
      <th style="width:11%">DATA</th>
      <th style="width:15%">AUTORE</th>
      <th style="width:30%">NOME FILE</th>
      <th style="width:44%">DESCRIZIONE</th>
    </tr>
  </thead>
  <tbody>${tableRows}</tbody>
</table>
</body></html>`;

  openPrintWindow(html, 'elenco_documentazione_grafica');
}

// ── Reset sezione tavole ────────────────────────────────

function tavReset() {
  tavProj = null; tavFiles = []; tavCounter = 0;
  document.getElementById('tavSelCantiere').value = '';
  document.getElementById('tavCantInfoBar').style.display = 'none';
  document.getElementById('tavBtnToUpload').style.display = 'none';
  document.getElementById('tavPreviewGrid').innerHTML = '';
  document.getElementById('tavRows').innerHTML = '';
  document.getElementById('tavUploadActions').style.display = 'none';
  document.getElementById('tavProgressBar').style.width = '0%';
  document.getElementById('btnTavNext').disabled = false;
  tavGoToPhase('tav-phase-cantiere', 1);
}
