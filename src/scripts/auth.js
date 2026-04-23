// ══════════════════════════════════════
//   AUTENTICACIÓN MICROSOFT 365 + MSAL
// ══════════════════════════════════════

var CURRENT_USER = null;
var msalApp      = null;

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
    },
    cache: { cacheLocation: 'sessionStorage' }
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

async function getToken(scopes) {
  const accounts = msalApp.getAllAccounts();
  if(!accounts.length) throw new Error('No session');
  try {
    const r = await msalApp.acquireTokenSilent({ scopes, account: accounts[0] });
    return r.accessToken;
  } catch {
    const r = await msalApp.acquireTokenPopup({ scopes });
    return r.accessToken;
  }
}

async function spLogin() {
  if(!initMsalApp()) throw new Error('MSAL no disponible');
  await msalApp.handleRedirectPromise();
  let account = msalApp.getAllAccounts()[0];
  if(!account) {
    await msalApp.loginPopup({ scopes: ['User.Read'] });
    account = msalApp.getAllAccounts()[0];
  }
  const token = await getToken(['User.Read']);
  const res   = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,otherMails', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const profile = await res.json();
  const { email, role, directorGroup, candidates } = resolveUserIdentity(profile, account);
  CURRENT_USER = { email, name: profile.displayName, role, directorGroup };
  sessionStorage.setItem('forecast_user', JSON.stringify(CURRENT_USER));
  console.log('[AUTH]', email, role, candidates);
  return true;
}

const ROLES = {
  gerencia: [
    'juannovoa@provexpress.com.co',
    'oscar.beltran@provexpress.com.co',
    'rafaelnovoa@provexpress.com.co',
    'c.estrategica@provexpress.com.co',
    'maribel.virguez@provexpress.com.co',
    'especialista.preventa@provexpress.com.co',
    'preventa.software@provexpress.com.co',
  ],
  directores: {
    'juannovoa@provexpress.com.co':          'Juan David Novoa',
    'angelica.caballero@provexpress.com.co': 'Maria Angelica Caballero',
    'oscar.beltran@provexpress.com.co':      'Oscar Beltran',
    'miller.romero@provexpress.com.co':      'Miller Romero',
  }
};

// Mapea correo -> nombre de archivo Excel del ejecutivo
const EXECUTIVO_BY_EMAIL = {
  'dafne.ruiz@provexpress.com.co': 'Dafne Lizeth Ruiz',
  'diana.castro@provexpress.com.co': 'Diana Catalina Castro',
  'jessica.valencia@provexpress.com.co': 'Jessica Lorena Valencia',
  'jhonatan.acevedo@provexpress.com.co': 'Jhonatan Acevedo',
  'camilo.hernandez@provexpress.com.co': 'Jhonatan Camilo Hernández',
  'juan.velasquez@provexpress.com.co': 'Juan Camilo Velásquez',
  'astrid.jimenez@provexpress.com.co': 'Leidy Astrid Jiménez',
  'maria.briceno@provexpress.com.co': 'María Paola Briceño',
  'yeison.urrego@provexpress.com.co': 'Yeison Urrego',
  'alejandra.velasquez@provexpress.com.co': 'Alejandra Velásquez',
  'angela.torres@provexpress.com.co': 'Angela Torres',
  'cesar.cespedes@provexpress.com.co': 'César Cespedes',
  'fernando.quinonez@provexpress.com.co': 'Fernando Alberto Quiñonez',
  'jenny.gonzalez@provexpress.com.co': 'Jenny González',
  'johanna.jaime@provexpress.com.co': 'Johanna Jaime Murcia',
  'juan.martinez@provexpress.com.co': 'Juan David Martínez',
  'mariela.ramirez@provexpress.com.co': 'Mariela Ramírez',
  'rosa.mendoza@provexpress.com.co': 'Rosa María Mendoza',
  'wilson.sanchez@provexpress.com.co': 'Wilson Fernando Sánchez',
  'tatiana.parra@provexpress.com.co': 'Angie Tatiana Parra',
  'claudia.triana@provexpress.com.co': 'Claudia Patricia Triana',
  'dilma.cuesta@provexpress.com.co': 'Dilma Cuesta',
  'andres.pena@provexpress.com.co': 'Freddy Andrés Peña',
  'paola.garcia@provexpress.com.co': 'Gina Paola García',
  'javier.cortes@provexpress.com.co': 'Javier Antonio Cortés',
  'julieth.galindo@provexpress.com.co': 'Juliet Milena Galindo Fino',
  'karen.carrillo@provexpress.com.co': 'Karent Carrillo',
  'lington.linares@provexpress.com.co': 'Lington Linares',
  'maria.cruz@provexpress.com.co': 'María Eugenia Cruz',
  'mario.reyes@provexpress.com.co': 'Mario Reyes',
  'daniel.galindo@provexpress.com.co': 'Daniel Galindo Girón',
  'dayana.chala@provexpress.com.co': 'Dayana Chala',
  'angelica.alvarez@provexpress.com.co': 'María Angélica Alvarez',
  'rosmira.rojas@provexpress.com.co': 'Rosmira Rojas',
  'yovanny.herrera@provexpress.com.co': 'Yovanny Herrera',
  'andrea.vargas@provexpress.com.co': 'Yurany Andrea Vargas',
};

const SALES_SUPPORT_BY_EMAIL = {
  'soporte.comercial5@provexpress.com.co': 'Isleni Yasmin Vasquez Pastrana',
  'soporte.comercial4@provexpress.com.co': 'Erika Gabriela Mieles Ortiz',
  'soporte.comercial6@provexpress.com.co': 'Nury Marcela Vargas Suarez',
  'soporte.comercial3@provexpress.com.co': 'Janira Alejandra Maldonado Prieto',
  'soporte.comercial@provexpress.com.co': 'Karen Cagua',
  'soporte.comercial2@provexpress.com.co': 'Alexandra Julieth Vargas Charris',
};

window.EXECUTIVO_BY_EMAIL = EXECUTIVO_BY_EMAIL;
window.SALES_SUPPORT_BY_EMAIL = SALES_SUPPORT_BY_EMAIL;

const SPECIAL_ROLE_IDENTITIES = [
  { email:'juannovoa@provexpress.com.co', name:'Juan David Novoa', role:'gerencia_director', directorGroup:'Juan David Novoa' },
  { email:'oscar.beltran@provexpress.com.co', name:'Oscar Beltran', role:'gerencia_director', directorGroup:'Oscar Beltran' },
];

function normalizeEmail(value) {
  return String(value || '').toLowerCase().trim();
}

function normalizeIdentityName(value) {
  return String(value || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function getSpecialRoleByName(name) {
  const target = normalizeIdentityName(name);
  if(!target) return null;
  return SPECIAL_ROLE_IDENTITIES.find(item => normalizeIdentityName(item.name) === target) || null;
}

function getIdentityCandidates(profile, account) {
  const candidates = [
    profile && profile.mail,
    profile && profile.userPrincipalName,
    account && account.username,
    ...((profile && profile.otherMails) || [])
  ].map(normalizeEmail).filter(Boolean);
  return [...new Set(candidates)];
}

function resolveUserIdentity(profile, account) {
  const candidates = getIdentityCandidates(profile, account);
  const execMap = window.EXECUTIVO_BY_EMAIL || EXECUTIVO_BY_EMAIL || {};

  for(const candidate of candidates) {
    const roleInfo = getUserRole(candidate);
    if(roleInfo.role !== 'ejecutivo') {
      return { email: candidate, role: roleInfo.role, directorGroup: roleInfo.directorGroup, candidates };
    }
  }

  for(const candidate of candidates) {
    if(execMap[candidate]) {
      const roleInfo = getUserRole(candidate);
      return { email: candidate, role: roleInfo.role, directorGroup: roleInfo.directorGroup, candidates };
    }
  }

  const specialRole = getSpecialRoleByName(profile && profile.displayName);
  if(specialRole) {
    const email = candidates[0] || specialRole.email;
    return { email, role: specialRole.role, directorGroup: specialRole.directorGroup, candidates };
  }

  const email = candidates[0] || '';
  const roleInfo = getUserRole(email);
  return { email, role: roleInfo.role, directorGroup: roleInfo.directorGroup, candidates };
}

function getUserRole(email) {
  const e = (email||'').toLowerCase().trim();
  const isGerencia = ROLES.gerencia.includes(e);
  const isDirector = e in ROLES.directores;
  const isSalesSupport = e in SALES_SUPPORT_BY_EMAIL;
  const dirGroup   = ROLES.directores[e] || null;
  if(isGerencia && isDirector) return { role:'gerencia_director', directorGroup: dirGroup };
  if(isGerencia)  return { role:'gerencia',  directorGroup: null };
  if(isDirector)  return { role:'director',  directorGroup: dirGroup };
  if(isSalesSupport) return { role:'sales_support', directorGroup: null };
  return { role:'ejecutivo', directorGroup: null };
}

function showUserBadge() {
  if(!CURRENT_USER) return;
  const badge = document.getElementById('user-badge');
  if(badge) badge.style.display = 'flex';
  const av = document.getElementById('user-avatar');
  if(av) av.textContent = CURRENT_USER.name.split(' ').slice(0,2).map(w=>w[0]).join('');
  const nm = document.getElementById('user-name');
  if(nm) nm.textContent = CURRENT_USER.name.split(' ')[0];
  const rb = document.getElementById('user-role-badge');
  const roleLabels = {gerencia:'Gerencia',gerencia_director:'Gerencia · Director',director:'Director',ejecutivo:'Ejecutivo',sales_support:'Sales Support'};
  if(rb) rb.textContent = roleLabels[CURRENT_USER.role]||CURRENT_USER.role;
  // Mostrar botón de cambio de vista solo para especialista.preventa
  const gearBtn = document.getElementById('view-switcher-btn');
  if(gearBtn && CURRENT_USER.email === 'especialista.preventa@provexpress.com.co') {
    gearBtn.style.display = 'block';
  }
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
    name: viewName,
    _realName: realName,
    _realEmail: realEmail
  };
  applyRoleTabs();
  renderAll();
  // Update role badge
  const rb = document.getElementById('user-role-badge');
  const roleLabels = {gerencia:'Gerencia',director:'Director',ejecutivo:'Ejecutivo',sales_support:'Sales Support'};
  let tail = '';
  if(role === 'director' && directorGroup) tail = ' · '+directorGroup.split(' ')[0];
  if(role === 'ejecutivo' && viewName) tail = ' · '+viewName.split(' ')[0];
  if(role === 'sales_support' && viewName) tail = ' · '+viewName.split(' ')[0];
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
    ov.style.cssText='position:fixed;inset:0;z-index:9999;background:rgba(3,5,14,.92);backdrop-filter:blur(8px);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px';
    ov.innerHTML=`
      <div style="font-family:var(--font-display);font-size:13px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--corp-cyan)">Conectando...</div>
      <div id="load-status" style="font-size:12px;color:var(--text3);font-family:var(--font-body);min-width:320px;text-align:center"></div>
      <div style="background:var(--border);border-radius:6px;height:4px;overflow:hidden;width:320px">
        <div id="load-bar" style="height:100%;background:linear-gradient(90deg,var(--corp-blue),var(--corp-cyan));width:100%;animation:pulse 1.5s ease-in-out infinite;border-radius:6px"></div>
      </div>`;
    document.body.appendChild(ov);
  }
  ov.style.display='flex';
  updateLoadingStatus(msg);
}
function updateLoadingStatus(msg){ const el=document.getElementById('load-status'); if(el) el.textContent=msg; }
function hideLoadingOverlay(){ const ov=document.getElementById('load-overlay'); if(ov) ov.style.display='none'; }
