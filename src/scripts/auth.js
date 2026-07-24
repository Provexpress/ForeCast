// ══════════════════════════════════════
//   AUTENTICACIÓN MICROSOFT 365 + MSAL
// ══════════════════════════════════════

var CURRENT_USER = null;
var msalApp      = null;
var msalRedirectPromise = null;
const AUTH_REDIRECT_ERROR_CODE = 'auth_redirect_started';
const AUTH_LOGIN_SCOPES = ['User.Read'];

function getAuthRedirectUri() {
  const origin = window.location.origin || '';
  let path = window.location.pathname || '/';
  path = path.split('?')[0].split('#')[0];
  if(path.endsWith('/index.html')) path = path.slice(0, -'/index.html'.length) || '/';
  if(!path.endsWith('/')) path += '/';
  return origin + path;
}

var AZURE_CONFIG = {
  clientId:  '4a2b9726-2736-4f72-9e7e-c64cfdc80253',
  tenantId:  'e6805558-f5bb-444c-8af2-5f3a4d6dd3fc',
  redirectUri: getAuthRedirectUri(),
  siteUrl:   'https://provexpress.sharepoint.com/sites/ProvexpressIntranet',
  driveBase: 'Documentos compartidos/COMERCIAL/FORECAST 2026',
};

function initMsalApp() {
  if(typeof msal === 'undefined') return false;
  if(msalApp) return true;
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId:    AZURE_CONFIG.clientId,
      authority:   'https://login.microsoftonline.com/' + AZURE_CONFIG.tenantId,
      redirectUri: AZURE_CONFIG.redirectUri,
      postLogoutRedirectUri: AZURE_CONFIG.redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
  });
  return true;
}

function loadMSAL(callback) {
  if(typeof msal !== 'undefined') { callback(); return; }
  const s = document.createElement('script');
  s.src = 'https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js';
  s.onload = () => callback();
  s.onerror = () => {
    const s2 = document.createElement('script');
    s2.src = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js';
    s2.onload = () => callback();
    s2.onerror = () => { console.error('[MSAL] Failed'); callback(); };
    document.head.appendChild(s2);
  };
  document.head.appendChild(s);
}

function getActiveMsalAccount() {
  return (msalApp.getActiveAccount && msalApp.getActiveAccount()) || msalApp.getAllAccounts()[0] || null;
}

function setActiveMsalAccount(account) {
  if(account && msalApp.setActiveAccount) msalApp.setActiveAccount(account);
}

async function handleMsalRedirectOnce() {
  if(!msalRedirectPromise) {
    msalRedirectPromise = msalApp.handleRedirectPromise().then(result => {
      if(result && result.account) setActiveMsalAccount(result.account);
      const account = getActiveMsalAccount();
      if(account) setActiveMsalAccount(account);
      return result;
    }).catch(error => {
      msalRedirectPromise = null;
      throw error;
    });
  }
  return msalRedirectPromise;
}

function createAuthRedirectError() {
  const error = new Error('Redirigiendo a Microsoft 365...');
  error.code = AUTH_REDIRECT_ERROR_CODE;
  return error;
}

function isAuthRedirectInProgress(error) {
  return error && error.code === AUTH_REDIRECT_ERROR_CODE;
}

async function startAuthRedirect(action, request) {
  updateLoadingStatus('Redirigiendo a Microsoft 365...');
  if(action === 'token') {
    await msalApp.acquireTokenRedirect(request);
  } else {
    await msalApp.loginRedirect(request);
  }
  throw createAuthRedirectError();
}

async function getToken(scopes) {
  await handleMsalRedirectOnce();
  const account = getActiveMsalAccount();
  if(!account) {
    await startAuthRedirect('login', { scopes: AUTH_LOGIN_SCOPES, prompt: 'select_account' });
  }
  try {
    const r = await msalApp.acquireTokenSilent({ scopes, account });
    return r.accessToken;
  } catch(error) {
    await startAuthRedirect('token', { scopes, account });
  }
}

async function spLogin() {
  if(!initMsalApp()) throw new Error('MSAL no disponible');
  await handleMsalRedirectOnce();
  let account = getActiveMsalAccount();
  if(!account) {
    await startAuthRedirect('login', { scopes: AUTH_LOGIN_SCOPES, prompt: 'select_account' });
  }
  const token = await getToken(AUTH_LOGIN_SCOPES);
  const res   = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,otherMails', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const profile = await res.json();
  const identity = resolveUserIdentity(profile, account);
  const { email, role, directorGroup, group, supportScope, supportedExecutives, candidates } = identity;
  CURRENT_USER = {
    email,
    name: identity.name || profile.displayName || email || 'Usuario',
    role,
    directorGroup,
    group,
    supportScope,
    supportedExecutives,
    noPermissions: !role
  };
  sessionStorage.setItem('forecast_user', JSON.stringify(CURRENT_USER));
  console.log('[AUTH]', email, role, candidates);
  if(!role) {
    showNoPermissionScreen(email);
    return false;
  }
  return true;
}

function getForecastStructure(){
  return window.FORECAST_STRUCTURE || {};
}

function normalizeAuthEmail(value) {
  const structure = getForecastStructure();
  return structure.normalizeEmail
    ? structure.normalizeEmail(value)
    : String(value || '').toLowerCase().trim();
}

function getIdentityCandidates(profile, account) {
  const candidates = [
    profile && profile.mail,
    profile && profile.userPrincipalName,
    account && account.username,
    ...((profile && profile.otherMails) || [])
  ].map(normalizeAuthEmail).filter(Boolean);
  return [...new Set(candidates)];
}

function resolveUserIdentity(profile, account) {
  const candidates = getIdentityCandidates(profile, account);
  for(const candidate of candidates) {
    const roleInfo = getUserRole(candidate);
    if(roleInfo.role) {
      return {
        email: candidate,
        name: getConfiguredUserName(candidate, roleInfo.role) || profile && profile.displayName || '',
        role: roleInfo.role,
        directorGroup: roleInfo.directorGroup,
        group: roleInfo.group,
        supportScope: roleInfo.supportScope,
        supportedExecutives: roleInfo.supportedExecutives,
        candidates
      };
    }
  }

  const email = candidates[0] || '';
  return {
    email,
    name: profile && profile.displayName || email,
    role: null,
    directorGroup: null,
    group: null,
    supportScope: '',
    supportedExecutives: [],
    candidates
  };
}

function getUserRole(email) {
  const structure = getForecastStructure();
  const e = normalizeAuthEmail(email);
  const role = structure.getRoleByEmail ? structure.getRoleByEmail(e) : null;
  const group = structure.getGroupByEmail ? structure.getGroupByEmail(e) : null;
  const directorGroup = group && structure.getDirectorNameByGroup ? structure.getDirectorNameByGroup(group) : null;
  const supportScope = structure.getSupportScopeByEmail ? structure.getSupportScopeByEmail(e) : '';
  const supportedExecutives = structure.getEjecutivosBySupport ? structure.getEjecutivosBySupport(e) : [];
  return {
    role,
    directorGroup,
    group,
    supportScope,
    supportedExecutives
  };
}

function getConfiguredUserName(email, role){
  const structure = getForecastStructure();
  if((role === 'director' || role === 'gerencia_director') && structure.getDirectorByEmail) {
    const director = structure.getDirectorByEmail(email);
    return director && director.nombre || '';
  }
  if(role === 'ejecutivo' && structure.getExecutiveDisplayNameByEmail) {
    return structure.getExecutiveDisplayNameByEmail(email) || '';
  }
  if((role === 'sales_support' || role === 'sales_support_comercial') && structure.getSupportDisplayNameByEmail) {
    return structure.getSupportDisplayNameByEmail(email) || '';
  }
  return '';
}

function escapeAuthHtml(value){
  return String(value || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function showNoPermissionScreen(email){
  hideLoadingOverlay();
  showUserBadge();
  const nav = document.querySelector('nav');
  const main = document.querySelector('main');
  if(nav) nav.style.display = 'none';
  if(main) main.style.display = 'none';
  let panel = document.getElementById('no-permission-panel');
  if(!panel) {
    panel = document.createElement('div');
    panel.id = 'no-permission-panel';
    panel.className = 'auth-loading-overlay';
    panel.innerHTML = `
      <section class="auth-loading-card" aria-label="Sin permisos">
        <div class="auth-loading-brand">
          <img src="src/Logo.webp" alt="Provexpress">
          <span>Forecast 2026</span>
        </div>
        <div class="auth-loading-pill">Acceso restringido</div>
        <h2>Sin permisos</h2>
        <p>Tu correo corporativo no esta configurado en la estructura comercial 2026.</p>
        <div class="auth-loading-status">${escapeAuthHtml(email || 'Usuario sin correo detectado')}</div>
      </section>`;
    document.body.appendChild(panel);
  }
  panel.style.display = 'flex';
}

function showUserBadge() {
  if(!CURRENT_USER) return;
  const badge = document.getElementById('user-badge');
  if(badge) badge.style.display = 'flex';
  const displayName = CURRENT_USER.name || CURRENT_USER.email || 'Usuario';
  const av = document.getElementById('user-avatar');
  if(av) av.textContent = displayName.split(' ').slice(0,2).map(w=>w[0]).join('');
  const nm = document.getElementById('user-name');
  if(nm) nm.textContent = displayName.split(' ')[0];
  const rb = document.getElementById('user-role-badge');
  const roleLabels = {
    gerencia: 'Gerencia',
    gerencia_director: 'Gerencia · Director',
    director: 'Director',
    ejecutivo: 'Ejecutivo',
    sales_support: 'Sales Support',
    sales_support_comercial: 'Sales Support · Comercial'
  };
  let roleText = roleLabels[CURRENT_USER.role] || CURRENT_USER.role || 'Sin permisos';
  if(CURRENT_USER.role === 'sales_support' && CURRENT_USER.supportScope === 'unit') roleText = 'Sales Support · Unidad';
  if(rb) rb.textContent = roleText;
  const gearBtn = document.getElementById('view-switcher-btn');
  if(gearBtn && (CURRENT_USER.role === 'gerencia' || CURRENT_USER.role === 'gerencia_director')) {
    renderViewPanelOptions();
    gearBtn.style.display = 'block';
  }
}

function renderViewPanelOptions(){
  const host = document.getElementById('view-panel-options');
  const structure = getForecastStructure();
  if(!host || !structure.getAllDirectorNames) return;
  const directorNames = structure.getAllDirectorNames();
  const directorButtons = directorNames.map(name =>
    `<button onclick="switchView(this,'director','${name.replace(/'/g, "\\'")}')" class="view-opt-btn" data-view="dir-${name.replace(/\s+/g, '-').toLowerCase()}">Director - ${name}</button>`
  ).join('');
  const executiveEntries = structure.getAllExecutiveEmails && structure.getExecutiveByEmail
    ? structure.getAllExecutiveEmails()
      .map(email => structure.getExecutiveByEmail(email))
      .filter(Boolean)
      .map(executive => ({
        name: executive.nombre,
        director: structure.getDirectorNameByGroup
          ? structure.getDirectorNameByGroup(executive.grupo)
          : ''
      }))
      .sort((a, b) =>
        a.director.localeCompare(b.director, 'es') ||
        a.name.localeCompare(b.name, 'es')
      )
    : [];
  const executiveButtons = executiveEntries.map(executive => `
    <button
      type="button"
      onclick="switchExecutiveView(this)"
      class="view-opt-btn view-opt-executive"
      data-executive-name="${escapeAuthHtml(executive.name)}"
      data-director-name="${escapeAuthHtml(executive.director)}"
      title="Ver la aplicación como ${escapeAuthHtml(executive.name)}">
      <span>${escapeAuthHtml(executive.name)}</span>
      <small>${escapeAuthHtml(executive.director || 'Sin director')}</small>
    </button>`
  ).join('');
  host.innerHTML = `
    <button onclick="switchView(this,'gerencia')" class="view-opt-btn" data-view="gerencia">Gerencia General</button>
    ${directorButtons}
    ${executiveButtons ? `
      <div class="view-opt-section-title">Vista como comercial</div>
      <div class="view-opt-executive-list">${executiveButtons}</div>
    ` : ''}`;
}

function switchExecutiveView(buttonEl){
  if(!buttonEl) return;
  const executiveName = buttonEl.dataset.executiveName || '';
  const directorName = buttonEl.dataset.directorName || '';
  if(!executiveName) return;
  switchView(buttonEl, 'ejecutivo', directorName, executiveName);
}

function toggleViewPanel() {
  const panel = document.getElementById('view-panel');
  if(!panel) return;
  panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}

// Close view panel clicking outside
document.addEventListener('click', e => {
  const panel = document.getElementById('view-panel');
  const btn   = document.getElementById('view-switcher-btn');
  if(panel && btn && !panel.contains(e.target) && !btn.contains(e.target)) {
    panel.style.display = 'none';
  }
});

function switchView(buttonEl, role, directorGroup, nameOverride) {
  // Override CURRENT_USER view without changing real identity
  const prev = CURRENT_USER;
  const realName = prev._realName || prev.name;
  const realEmail = prev._realEmail || prev.email;
  const viewName = (nameOverride && String(nameOverride).trim()) || realName;
  CURRENT_USER = {
    ...prev,
    role,
    directorGroup: directorGroup||null,
    group: role === 'director' ? null : prev.group,
    name: viewName,
    _realName: realName,
    _realEmail: realEmail
  };
  applyRoleTabs();
  // Update role badge
  const rb = document.getElementById('user-role-badge');
  const roleLabels = {gerencia:'Gerencia',director:'Director',ejecutivo:'Ejecutivo',sales_support:'Sales Support',sales_support_comercial:'Sales Support · Comercial'};
  let tail = '';
  if(role === 'director' && directorGroup) tail = ' · '+directorGroup.split(' ')[0];
  if(role === 'ejecutivo' && viewName) tail = ' · '+viewName.split(' ')[0];
  if((role === 'sales_support' || role === 'sales_support_comercial') && viewName) tail = ' · '+viewName.split(' ')[0];
  if(rb) rb.textContent = (roleLabels[role]||role) + tail + ' ⚙';
  // Highlight active button
  document.querySelectorAll('.view-opt-btn').forEach(b => b.classList.remove('active'));
  if(buttonEl) buttonEl.classList.add('active');
  // Close panel
  const panel = document.getElementById('view-panel');
  if(panel) panel.style.display = 'none';
}

// ── overlay helpers ──────────────────────────
function showLoadingOverlay(msg) {
  let ov = document.getElementById('load-overlay');
  if(!ov){
    ov = document.createElement('div');
    ov.id = 'load-overlay';
    ov.className = 'auth-loading-overlay';
    ov.innerHTML=`
      <section class="auth-loading-card" aria-label="Validando acceso">
        <div class="auth-loading-brand">
          <img src="src/Logo.webp" alt="Provexpress">
          <span>Forecast 2026</span>
        </div>
        <div class="auth-loading-pill">Acceso corporativo</div>
        <h2>Validando acceso</h2>
        <p>Estamos conectando con Microsoft 365 y preparando la informacion comercial.</p>
        <div id="load-status" class="auth-loading-status"></div>
        <div class="auth-loading-track">
          <div id="load-bar" class="auth-loading-bar"></div>
        </div>
      </section>`;
    document.body.appendChild(ov);
  }
  ov.style.display='flex';
  updateLoadingStatus(msg);
}
function updateLoadingStatus(msg){ const el=document.getElementById('load-status'); if(el) el.textContent=msg; }
function hideLoadingOverlay(){ const ov=document.getElementById('load-overlay'); if(ov) ov.style.display='none'; }
