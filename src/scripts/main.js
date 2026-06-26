/* ══════════════════════════════════════
   DATA & STATE
══════════════════════════════════════ */
let ALL_DATA = [];
let SALES_DATA = [];
let SALES_PENDING_DATA = [];
let PREVENTA_DATA = [];
let TRM = 4150;
let TRM_READY = false;
let SELECTED_EXEC_BY_DIR = {};
let RECORD_SEQ = 0;
let NEGOCIO_DETAIL_STATE = null;
let MARCA_LINEA_DETAIL_STATE = null;
let EJECUTIVO_BRAND_FOCUS = null;
let LOADED_SALES_BY_SUPPORT = {};
let LOADED_PREVENTA_FILES = [];
let SALES_VIEW_MODE = 'reporte';
let SALES_PENDING_VALUE_ONLY = false;
let FINANCE_STATE = {
  isLoading: false,
  data: null,
  error: null,
  period: '',
  endDate: '',
};
let FORECAST_CONNECTIONS_LIST_ID = null;
let FORECAST_CONNECTIONS = { byEmail: {}, byName: {} };
let GERENCIA_CROSSFILTERS = { estado:'', director:'', ejecutivo:'', linea:'' };
let EXCELJS_EXPORT_LOAD_PROMISE = null;
let VISUAL_CROSSFILTERS = {
  ejecutivo: { linea:'' },
  sales: { soporta:'' },
  preventa: { linea:'' },
  marcas: { marca:'', linea:'' }
};

const GERENCIA_ESTADO_STEP = 20;
const GERENCIA_ESTADO_LIMITS = {
  GANADA: 30,
  PENDIENTE: 30,
  PERDIDA: 30,
  APLAZADO: 30
};
const GERENCIA_ESTADOS = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'];
const GERENCIA_ESTADO_COLORS = {
  GANADA: '#0DBF82',
  PENDIENTE: '#F0A020',
  PERDIDA: '#2D4FD6',
  APLAZADO: '#E84040'
};
const GERENCIA_COMITE_THRESHOLD_COP = 1000000000;
const GERENCIA_EXCEL_DETAIL_COLS = 21;
const DIVISAS_DETAIL_STEP = 10;
const DIVISAS_DETAIL_LIMITS = { COP: 10, USD: 10 };
const TOP_BAR_LIMIT = 15;
const MARCAS_BAR_STEP = 10;
const MARCAS_BAR_INITIAL = 10;
let MARCAS_BAR_LIMIT = MARCAS_BAR_INITIAL;

const TRM_CACHE_KEY = 'trm_last';
const FINANCE_GOAL_KEY_PREFIX = 'finance_sales_goal_';
const FINANCE_PUBLIC_CACHE_URL = 'https://tableros-area-financiera.vercel.app/_cache/ventas-pbi.json';
const FINANCE_CATEGORY_ORDER = ['Servicios','Renta de Equipos','PAC','Suministros','Tecnología','Licenciamiento'];
const FINANCE_CATEGORY_COLORS = ['#0DBF82','#2ABFDF','#F0A020','#8B5FC8','#2D4FD6','#E84040','#40C8F0','#E040A0'];
const FORECAST_CONNECTIONS_LIST_NAME = 'ForecastConexiones';
const PREVENTA_FOLDER_NAME = 'Grupo preventa';
const FORECAST_BASE_FALLBACK = 'COMERCIAL/FORECAST 2026';

function loadCachedTRM(){
  try {
    const v = parseFloat(localStorage.getItem(TRM_CACHE_KEY));
    if(v && v > 100){
      TRM = v;
      TRM_READY = true;
      const inp = document.getElementById('trm-input');
      if(inp) inp.value = Number(TRM).toFixed(2);
    }
  } catch(_) {}
}
function cacheTRM(val){
  try {
    if(val && val > 100){
      localStorage.setItem(TRM_CACHE_KEY, String(val));
    }
  } catch(_) {}
}

loadCachedTRM();

const APP_THEME_KEY = 'forecast_theme';

function applyAppTheme(theme){
  const selected = theme === 'light' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', selected);
  try { localStorage.setItem(APP_THEME_KEY, selected); } catch(_) {}
  const btn = document.getElementById('theme-toggle-btn');
  if(btn) {
    btn.textContent = selected === 'light' ? '☀' : '☾';
    btn.title = selected === 'light' ? 'Cambiar a modo oscuro' : 'Cambiar a modo claro';
    btn.setAttribute('aria-label', btn.title);
  }
}

function toggleAppTheme(){
  const current = document.documentElement.getAttribute('data-theme') === 'light' ? 'light' : 'dark';
  applyAppTheme(current === 'light' ? 'dark' : 'light');
}

function initAppTheme(){
  let saved = 'dark';
  try { saved = localStorage.getItem(APP_THEME_KEY) || 'dark'; } catch(_) {}
  applyAppTheme(saved);
}

initAppTheme();

async function fetchTRM() {
  const setTRM = t => {
    TRM = t;
    TRM_READY = true;
    const inp = document.getElementById('trm-input');
    if(inp) inp.value = Number(TRM).toFixed(2);
    if(ALL_DATA.length || SALES_DATA.length || SALES_PENDING_DATA.length || PREVENTA_DATA.length) renderVisiblePage();
    cacheTRM(TRM);
    console.log('[TRM]', TRM);
  };

  // Source 1: datos.gov.co — API oficial Superfinanciera (CORS abierto, sin proxy)
  try {
    // Misma lógica que Excel: vigenciadesde <= hoy AND vigenciahasta >= hoy
    const hoy = new Date(Date.now() - 5*60*60*1000).toISOString().substring(0,10);
    const url = 'https://www.datos.gov.co/resource/32sa-8pi3.json'
      + '?$select=valor,vigenciadesde,vigenciahasta'
      + `&$where=vigenciadesde <= '${hoy}' and vigenciahasta >= '${hoy}'`;
    const r1 = await fetch(url, { cache: 'no-store' });
    if(r1.ok) {
      const d1 = await r1.json();
      if(d1 && d1[0] && d1[0].valor) {
        const t1 = parseFloat(d1[0].valor);
        if(t1 > 100) { setTRM(t1); return; }
      }
    }
  } catch(e1) { console.warn('[TRM] datos.gov.co failed', e1.message); }

  // Source 2: local Vercel function (fallback on deployed environment)
  try {
    const r2 = await fetch('/api/trm', { cache: 'no-store' });
    if(r2.ok) {
      const d2 = await r2.json();
      if(d2.ok && d2.trm > 100) { setTRM(d2.trm); return; }
    }
  } catch(e2) { console.warn('[TRM] vercel failed', e2.message); }

  // Source 3: open.er-api.com (fallback)
  try {
    const r3 = await fetch('https://open.er-api.com/v6/latest/USD', { cache: 'no-store' });
    if(r3.ok) {
      const d3 = await r3.json();
      const t3 = d3.rates && d3.rates.COP;
      if(t3 > 100) { setTRM(t3); return; }
    }
  } catch(e3) { console.warn('[TRM] open.er-api failed', e3.message); }

  console.warn('[TRM] all sources failed, keeping current TRM', TRM);
}

const COLORS = [
  '#2D4FD6','#8B5FC8','#2ABFDF','#0DBF82','#F0A020',
  '#E84040','#E040A0','#40C8F0','#F06040','#1B2B8C'
];
const FORECAST_YEAR = '2026';
const MONTH_LABELS_SHORT = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const MONTH_LABELS_LONG = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

/* Format helpers */
function fmtCOP(v){
  if(v===null||v===undefined) return '—';
  return '$ ' + Math.round(v).toLocaleString('es-CO');
}
function fmtUSD(v){
  if(v===null||v===undefined) return '—';
  return 'USD ' + Number(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function fmtUSDCompact(v){
  if(v===null||v===undefined||isNaN(v)) return '—';
  const n=Number(v);
  const abs=Math.abs(n);
  if(abs>=1e9) return 'USD '+(n/1e9).toFixed(2)+' B';
  if(abs>=1e6) return 'USD '+(n/1e6).toFixed(1)+' M';
  if(abs>=1e3) return 'USD '+(n/1e3).toFixed(0)+' K';
  return 'USD '+n.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:2});
}
function kpiUSD(v){
  return `<div class="kpi-val kpi-val-usd" title="${escAttr(fmtUSD(v))}">${escHtml(fmtUSDCompact(v))}</div>`;
}
function fmtTRM(v){
  if(v===null||v===undefined||isNaN(v)) return '—';
  return Number(v).toLocaleString('es-CO',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function fmtTRMDisplay(v){
  return TRM_READY ? fmtTRM(v) : '—';
}
function graphJsonHeaders(token, extra = {}) {
  return {
    Authorization: 'Bearer ' + token,
    Accept: 'application/json',
    'Content-Type': 'application/json',
    ...extra
  };
}
function getRoleLabel(role) {
  return ({
    gerencia: 'Gerencia',
    gerencia_director: 'Gerencia · Director',
    director: 'Director',
    ejecutivo: 'Ejecutivo',
    sales_support: 'Sales Support'
  })[role] || role || '';
}
function formatLastConnection(value) {
  if(!value) return '';
  const date = new Date(value);
  if(Number.isNaN(date.getTime())) return '';
  const localDate = new Date(date.toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const target = new Date(localDate.getFullYear(), localDate.getMonth(), localDate.getDate()).getTime();
  const time = localDate.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });
  const diffDays = Math.round((today - target) / 86400000);
  if(diffDays === 0) return 'Hoy ' + time;
  if(diffDays === 1) return 'Ayer ' + time;
  return localDate.toLocaleDateString('es-CO', { day: '2-digit', month: 'short' }) + ' ' + time;
}
function renderLastConnection(name) {
  const fileInfo = getLoadedExcelFileMetaForExecutive(name);
  const text = fileInfo && fileInfo.lastModifiedDateTime ? formatLastConnection(fileInfo.lastModifiedDateTime) : '';
  return `<div class="last-connection ${text ? '' : 'empty'}">
    <span></span>${text ? 'Ult. Excel: ' + escHtml(text) : 'Sin fecha del Excel'}
  </div>`;
}

function formatLastConnectionFull(value){
  if(!value) return 'Sin fecha registrada';
  const date = new Date(value);
  if(Number.isNaN(date.getTime())) return 'Sin fecha registrada';
  return date.toLocaleString('es-CO', {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function getExcelFileStatusInfo(value){
  if(!value) return 'Sin fecha';
  const date = new Date(value);
  if(Number.isNaN(date.getTime())) return 'Sin fecha';
  const diffHours = (Date.now() - date.getTime()) / 3600000;
  if(diffHours <= 24) return 'Actualizado hoy';
  if(diffHours <= 168) return 'Actualizado esta semana';
  return 'Sin actualizacion reciente';
}

function getDriveItemModifier(item){
  const by = item && item.lastModifiedBy || {};
  const user = by.user || by.application || {};
  return {
    name: cleanDisplayText(user.displayName, ''),
    email: cleanDisplayText(user.email || user.userPrincipalName, '')
  };
}

function buildLoadedExcelFileMeta(item, dirName){
  const modifier = getDriveItemModifier(item || {});
  return {
    name: item && item.name || '',
    dir: dirName || '',
    lastModifiedDateTime: item && item.lastModifiedDateTime || '',
    lastModifiedByName: modifier.name,
    lastModifiedByEmail: modifier.email,
    webUrl: item && item.webUrl || '',
    size: Number(item && item.size || 0)
  };
}

function getLoadedExcelFileMetaForExecutive(name, directorHint){
  const target = cleanDisplayText(name, '');
  const directorNorm = normalizePersonName(directorHint || '');
  const allFiles = Object.entries(LOADED_FILES_BY_DIR || {}).flatMap(([director, files]) =>
    (files || []).map(file => ({ ...file, dir: file.dir || director }))
  );
  const direct = allFiles.find(file => {
    if(directorNorm && normalizePersonName(file.dir) !== directorNorm) return false;
    return namesMatch(String(file.name || '').replace(/\.(xlsx|xls)$/i, ''), target);
  });
  if(direct) return direct;
  return allFiles.find(file => namesMatch(String(file.name || '').replace(/\.(xlsx|xls)$/i, ''), target)) || null;
}

function getExecutiveExcelMetadataRows(executives, sourceRows, opts){
  const options = opts || {};
  const rows = Array.isArray(sourceRows) ? sourceRows : [];
  return (executives || []).filter(Boolean).map(name => {
    const rowMatch = rows.find(row => namesMatch(row['COMERCIAL'] || '', name));
    const fileInfo = getLoadedExcelFileMetaForExecutive(name, options.director) || {};
    const director = cleanDisplayText(options.director || fileInfo.dir || rowMatch && rowMatch['DIRECTOR'], 'Sin director');
    const file = cleanDisplayText(fileInfo.name || rowMatch && rowMatch.__SOURCE_FILE, cleanDisplayText(name, 'Ejecutivo') + '.xlsx');
    const personRows = rows.filter(row => namesMatch(row['COMERCIAL'] || '', name));
    return {
      name,
      director,
      file,
      rawDate: fileInfo.lastModifiedDateTime || '',
      lastModified: formatLastConnectionFull(fileInfo.lastModifiedDateTime),
      lastModifiedBy: cleanDisplayText(fileInfo.lastModifiedByName, 'Sin dato'),
      lastModifiedByEmail: cleanDisplayText(fileInfo.lastModifiedByEmail, ''),
      webUrl: cleanDisplayText(fileInfo.webUrl, ''),
      sizeKb: fileInfo.size ? Math.round(fileInfo.size / 1024) : 0,
      status: getExcelFileStatusInfo(fileInfo.lastModifiedDateTime),
      negocios: personRows.length
    };
  }).sort((a,b) => {
    const dateA = a.rawDate ? new Date(a.rawDate).getTime() : 0;
    const dateB = b.rawDate ? new Date(b.rawDate).getTime() : 0;
    if(dateA !== dateB) return dateB - dateA;
    return a.name.localeCompare(b.name, 'es', { sensitivity:'base' });
  });
}

function getExcelMetadataDateKey(value){
  if(!value) return '';
  const date = new Date(value);
  if(Number.isNaN(date.getTime())) return '';
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(date).reduce((acc, part) => {
    if(part.type !== 'literal') acc[part.type] = part.value;
    return acc;
  }, {});
  return parts.year && parts.month && parts.day ? `${parts.year}-${parts.month}-${parts.day}` : '';
}

function formatExcelMetadataDateLabel(dateKey){
  if(!dateKey) return 'Sin fecha';
  const parts = String(dateKey).split('-').map(Number);
  if(parts.length !== 3 || parts.some(n => !n)) return dateKey;
  return new Date(parts[0], parts[1] - 1, parts[2]).toLocaleDateString('es-CO', {
    day: '2-digit',
    month: 'short'
  });
}

function getExcelMetadataDateBuckets(metadataRows){
  const buckets = new Map();
  (metadataRows || []).forEach(item => {
    const key = getExcelMetadataDateKey(item.rawDate);
    if(!key) return;
    const current = buckets.get(key) || { key, count: 0, names: [], files: [] };
    current.count += 1;
    current.names.push(item.name);
    current.files.push(item.file);
    buckets.set(key, current);
  });
  return [...buckets.values()].sort((a,b) => a.key.localeCompare(b.key));
}

function buildExcelConnectionExportButton(scope, label){
  const id = scope === 'director' ? 'btn-export-director-excel-meta' : 'btn-export-ejecutivo-excel-meta';
  return `<button id="${id}" class="export-excel-btn metadata-export-btn" type="button" onclick="${escAttr(jsCall('downloadExecutiveExcelMetadataReport', scope))}" title="Descargar ultima actualizacion de Excel">
    <span class="export-excel-icon" aria-hidden="true">↓</span>
    <span>${escHtml(label || 'Excel')}</span>
  </button>`;
}

function getExecutiveExcelMetadataReportContext(scope){
  const reportScope = scope === 'ejecutivo' ? 'ejecutivo' : 'director';
  if(reportScope === 'ejecutivo') {
    const rows = getVisibleData();
    const execsFromData = [...new Set(rows.map(row => cleanDisplayText(row['COMERCIAL'], '')).filter(Boolean))];
    const execsFromFiles = Object.values(LOADED_FILES_BY_DIR || {}).flat()
      .map(file => String(file.name || '').replace(/\.(xlsx|xls)$/i, '').trim())
      .filter(Boolean);
    const executives = [...new Set([...execsFromData, ...execsFromFiles])].sort();
    return {
      scope: reportScope,
      title: 'Ultima actualizacion del Excel - Ejecutivos',
      subtitle: 'Global ejecutivos',
      executives,
      sourceRows: rows,
      director: ''
    };
  }

  const dir = document.getElementById('sel-director') ? document.getElementById('sel-director').value : '';
  const rows = getVisibleData().filter(row => cleanDisplayText(row['DIRECTOR'], '') === dir);
  const execsWithData = [...new Set(rows.map(row => cleanDisplayText(row['COMERCIAL'], '')).filter(Boolean))];
  const execsFromFiles = (LOADED_FILES_BY_DIR[dir] || []).map(file => String(file.name || '').replace(/\.(xlsx|xls)$/i, '').trim()).filter(Boolean);
  const executives = [...new Set([...execsWithData, ...execsFromFiles])].sort();
  return {
    scope: reportScope,
    title: 'Ultima actualizacion del Excel - Director',
    subtitle: dir || 'Director',
    executives,
    sourceRows: rows,
    director: dir
  };
}

function styleExcelMetadataKpis(ws, rows){
  const datedRows = (rows || []).filter(item => item.rawDate);
  const latest = datedRows[0] || null;
  const noDateCount = (rows || []).filter(item => !item.rawDate).length;
  const now = Date.now();
  const recentCount = datedRows.filter(item => {
    const time = new Date(item.rawDate).getTime();
    return !Number.isNaN(time) && (now - time) <= 7 * 24 * 3600000;
  }).length;
  const cards = [
    ['Archivos', rows.length, '#2D4FD6', '#,##0'],
    ['Actualizados en 7 dias', recentCount, '#0DBF82', '#,##0'],
    ['Mas reciente', latest ? new Date(latest.rawDate) : '', '#2ABFDF', 'yyyy-mm-dd hh:mm'],
    ['Sin fecha', noDateCount, '#F0A020', '#,##0']
  ];
  cards.forEach((card, idx) => {
    const startCol = idx * 4 + 1;
    const endCol = startCol + 3;
    ws.mergeCells(5, startCol, 5, endCol);
    ws.mergeCells(6, startCol, 6, endCol);
    const label = ws.getCell(5, startCol);
    const value = ws.getCell(6, startCol);
    label.value = card[0];
    value.value = card[1];
    value.numFmt = card[3];
    [label, value].forEach(cell => {
      cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
      cell.border = {
        top:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
        left:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
        bottom:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
        right:{style:'thin', color:{argb:excelArgb('#D7E0F0')}}
      };
      cell.alignment = { vertical:'middle', horizontal:'center' };
    });
    label.font = { name:'Aptos', size:9, bold:true, color:{ argb: excelArgb('#677592') } };
    value.font = { name:'Aptos Display', size:13, bold:true, color:{ argb: excelArgb(card[2]) } };
  });
  ws.getRow(5).height = 20;
  ws.getRow(6).height = 24;
}

function addExecutiveExcelMetadataChartSheet(workbook, context, rows){
  const ws = workbook.addWorksheet('Grafica fechas');
  const lastCol = 16;
  setupExcelWorksheet(ws, lastCol, 12);
  styleExcelTitle(
    ws,
    'Grafica de actualizaciones de Excel',
    `Generado: ${getBogotaTimestampLabel()} | ${context.subtitle || 'Sin filtro'}`,
    'Fuente: SharePoint',
    lastCol
  );
  ws.getCell('A3').value = 'Fuente: fecha de ultima modificacion del archivo Excel en SharePoint (lastModifiedDateTime). No corresponde al ingreso al dashboard.';
  styleExcelMetadataKpis(ws, rows);
  styleExcelSectionLabel(ws, 8, 'Como leer esta hoja', lastCol);
  ws.mergeCells(`A9:${excelColumnLetter(lastCol)}9`);
  const helpCell = ws.getCell('A9');
  helpCell.value = 'Cada barra muestra cuantos archivos Excel fueron modificados en esa fecha. La escala es relativa al dia con mas archivos actualizados.';
  helpCell.font = { name:'Aptos', size:10, color:{ argb: excelArgb('#677592') } };
  helpCell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
  helpCell.alignment = { vertical:'middle', horizontal:'left', wrapText:true };
  ws.getRow(9).height = 24;

  styleExcelSectionLabel(ws, 11, 'Distribucion por fecha', lastCol);

  const barStartCol = 4;
  const barEndCol = lastCol;
  const barSegments = barEndCol - barStartCol + 1;
  [16,12,14].forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });
  for(let col = barStartCol; col <= barEndCol; col++) ws.getColumn(col).width = 4;

  const headerRowNumber = 12;
  const dataStartRow = headerRowNumber + 1;
  const header = ws.getRow(headerRowNumber);
  ['Fecha','Archivos','Participacion'].forEach((label, idx) => {
    const cell = ws.getCell(headerRowNumber, idx + 1);
    cell.value = label;
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#1B2B8C') } };
    cell.font = { name:'Aptos', size:9, bold:true, color:{ argb:'FFFFFFFF' } };
    cell.alignment = { vertical:'middle', horizontal:'center' };
    cell.border = {
      top:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
      left:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
      bottom:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
      right:{style:'thin', color:{argb:excelArgb('#D7E0F0')}}
    };
  });
  ws.mergeCells(headerRowNumber, barStartCol, headerRowNumber, barEndCol);
  const chartHeader = ws.getCell(headerRowNumber, barStartCol);
  chartHeader.value = 'Barra de archivos actualizados';
  chartHeader.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#1B2B8C') } };
  chartHeader.font = { name:'Aptos', size:9, bold:true, color:{ argb:'FFFFFFFF' } };
  chartHeader.alignment = { vertical:'middle', horizontal:'center' };
  chartHeader.border = {
    top:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
    left:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
    bottom:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
    right:{style:'thin', color:{argb:excelArgb('#D7E0F0')}}
  };
  header.height = 22;

  const buckets = getExcelMetadataDateBuckets(rows).slice(-14);
  const totalDated = buckets.reduce((sum,item) => sum + item.count, 0) || 1;
  const maxCount = Math.max(...buckets.map(item => item.count), 1);
  const borderColor = excelArgb('#D7E0F0');

  if(!buckets.length){
    ws.mergeCells(`A${dataStartRow}:${excelColumnLetter(lastCol)}${dataStartRow}`);
    const cell = ws.getCell(`A${dataStartRow}`);
    cell.value = 'No hay fechas de actualizacion disponibles en los metadatos de SharePoint.';
    cell.font = { name:'Aptos', size:11, italic:true, color:{ argb: excelArgb('#677592') } };
    cell.alignment = { vertical:'middle', horizontal:'center' };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
    ws.getRow(dataStartRow).height = 24;
    return;
  }

  buckets.forEach((item, idx) => {
    const rowNumber = dataStartRow + idx;
    const row = ws.getRow(rowNumber);
    row.height = 22;
    const fillColor = excelArgb('#2D4FD6');
    const segments = Math.max(1, Math.round((item.count / maxCount) * barSegments));
    ws.getCell(rowNumber, 1).value = formatExcelMetadataDateLabel(item.key);
    ws.getCell(rowNumber, 2).value = item.count;
    ws.getCell(rowNumber, 3).value = item.count / totalDated;
    ws.getCell(rowNumber, 2).numFmt = '#,##0';
    ws.getCell(rowNumber, 3).numFmt = '0.##%';

    for(let col = 1; col <= lastCol; col++){
      const cell = ws.getCell(rowNumber, col);
      cell.border = {
        top:{style:'thin', color:{argb:borderColor}},
        left:{style:'thin', color:{argb:borderColor}},
        bottom:{style:'thin', color:{argb:borderColor}},
        right:{style:'thin', color:{argb:borderColor}}
      };
      cell.font = { name:'Aptos', size:9, color:{ argb: excelArgb('#172033') } };
      cell.alignment = { vertical:'middle', horizontal: col >= 2 && col <= 3 ? 'right' : 'left' };
    }

    for(let col = barStartCol; col <= barEndCol; col++){
      const barCell = ws.getCell(rowNumber, col);
      const filled = col < barStartCol + segments;
      barCell.value = filled && col === barStartCol ? `${item.count}` : '';
      barCell.fill = {
        type:'pattern',
        pattern:'solid',
        fgColor:{ argb: filled ? fillColor : excelArgb('#F6F8FF') }
      };
      barCell.font = { name:'Aptos', size:8, bold:true, color:{ argb: filled ? 'FFFFFFFF' : excelArgb('#677592') } };
      barCell.alignment = { vertical:'middle', horizontal:'center' };
    }
  });

  const noteRow = dataStartRow + buckets.length + 2;
  ws.mergeCells(`A${noteRow}:${excelColumnLetter(lastCol)}${noteRow}`);
  const noteCell = ws.getCell(`A${noteRow}`);
  noteCell.value = 'El listado de ejecutivos y archivos por fecha esta en la hoja "Detalle por fecha".';
  noteCell.font = { name:'Aptos', size:10, italic:true, color:{ argb: excelArgb('#677592') } };
  noteCell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
  noteCell.alignment = { vertical:'middle', horizontal:'left' };
  ws.getRow(noteRow).height = 22;
}

function addExecutiveExcelMetadataDateDetailSheet(workbook, context, rows){
  const ws = workbook.addWorksheet('Detalle por fecha');
  const lastCol = 4;
  setupExcelWorksheet(ws, lastCol, 6);
  styleExcelTitle(
    ws,
    'Detalle de actualizaciones por fecha',
    `Generado: ${getBogotaTimestampLabel()} | ${context.subtitle || 'Sin filtro'}`,
    'Fuente: SharePoint lastModifiedDateTime',
    lastCol
  );
  [18,12,56,56].forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });

  const buckets = getExcelMetadataDateBuckets(rows).slice(-30);
  if(!buckets.length){
    ws.mergeCells(`A6:${excelColumnLetter(lastCol)}6`);
    const cell = ws.getCell('A6');
    cell.value = 'No hay fechas de actualizacion disponibles en los metadatos de SharePoint.';
    cell.font = { name:'Aptos', size:11, italic:true, color:{ argb: excelArgb('#677592') } };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
    cell.alignment = { vertical:'middle', horizontal:'center' };
    ws.getRow(6).height = 24;
    return;
  }

  const detailColumns = [
    { name:'Fecha', totalsRowLabel:'Total' },
    { name:'Archivos', totalsRowFunction:'sum' },
    { name:'Ejecutivos' },
    { name:'Archivos Excel' }
  ];
  const detailRows = buckets.map(item => [
    formatExcelMetadataDateLabel(item.key),
    item.count,
    excelSafeText(item.names.join(', ')),
    excelSafeText(item.files.join(', '))
  ]);
  addExcelTable(ws, {
    name: context.scope === 'ejecutivo' ? 'DetalleFechasEjecutivos' : 'DetalleFechasDirector',
    ref: 'A6',
    columns: detailColumns,
    rows: detailRows,
    theme: 'TableStyleMedium2'
  });
  styleExcelTable(ws, 6, detailRows.length, detailColumns.length, {
    wrapColumns: [3,4],
    rightColumns: [2]
  });
  for(let rowNumber = 7; rowNumber <= 7 + detailRows.length; rowNumber++){
    ws.getCell(rowNumber, 2).numFmt = '#,##0';
  }
}

function buildExecutiveExcelMetadataWorkbook(context){
  const rows = getExecutiveExcelMetadataRows(context.executives, context.sourceRows, { director: context.director });
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Forecast 2026 Provexpress';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.company = 'Provexpress';
  workbook.subject = 'Ultima actualizacion de Excel';
  workbook.title = context.title;
  workbook.calcProperties = { fullCalcOnLoad: true };

  addExecutiveExcelMetadataChartSheet(workbook, context, rows);
  addExecutiveExcelMetadataDateDetailSheet(workbook, context, rows);

  const ws = workbook.addWorksheet('Ultima actualizacion Excel');
  setupExcelWorksheet(ws, 10, 6);
  styleExcelTitle(
    ws,
    context.title,
    `Generado: ${getBogotaTimestampLabel()} | Registros: ${rows.length}`,
    context.subtitle || 'Sin filtro',
    10
  );
  [26,24,32,20,24,28,12,12,20,46].forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });
  addExcelTable(ws, {
    name: context.scope === 'ejecutivo' ? 'UltimaActualizacionEjecutivo' : 'UltimaActualizacionDirector',
    ref: 'A6',
    columns: [
      { name:'Ejecutivo', totalsRowLabel:'Total' },
      { name:'Director' },
      { name:'Archivo Excel' },
      { name:'Ultima actualizacion Excel' },
      { name:'Modificado por' },
      { name:'Correo modificador' },
      { name:'Negocios', totalsRowFunction:'sum' },
      { name:'Tamano KB', totalsRowFunction:'sum' },
      { name:'Estado' },
      { name:'Link SharePoint' }
    ],
    rows: rows.map(item => [
      excelSafeText(item.name),
      excelSafeText(item.director),
      excelSafeText(item.file),
      item.rawDate ? new Date(item.rawDate) : '',
      excelSafeText(item.lastModifiedBy),
      excelSafeText(item.lastModifiedByEmail),
      item.negocios,
      item.sizeKb,
      excelSafeText(item.status),
      excelSafeText(item.webUrl)
    ]),
    theme: 'TableStyleMedium4'
  });
  styleExcelTable(ws, 6, rows.length, 10, {
    wrapColumns: [3,10],
    rightColumns: [7,8]
  });
  for(let rowNumber = 7; rowNumber <= 7 + rows.length; rowNumber++){
    ws.getCell(rowNumber, 4).numFmt = 'yyyy-mm-dd hh:mm';
    ws.getCell(rowNumber, 7).numFmt = '#,##0';
    ws.getCell(rowNumber, 8).numFmt = '#,##0';
  }
  return workbook;
}

function setExecutiveExcelMetadataButtonBusy(scope, isBusy){
  const id = scope === 'ejecutivo' ? 'btn-export-ejecutivo-excel-meta' : 'btn-export-director-excel-meta';
  const btn = document.getElementById(id);
  if(!btn) return;
  btn.disabled = Boolean(isBusy);
  btn.innerHTML = isBusy
    ? '<span class="export-excel-icon" aria-hidden="true">...</span><span>Generando</span>'
    : '<span class="export-excel-icon" aria-hidden="true">↓</span><span>Excel</span>';
}

async function downloadExecutiveExcelMetadataReport(scope){
  const reportScope = scope === 'ejecutivo' ? 'ejecutivo' : 'director';
  const context = getExecutiveExcelMetadataReportContext(reportScope);
  if(!context.executives.length){
    alert('No hay ejecutivos para descargar en esta vista.');
    return;
  }

  setExecutiveExcelMetadataButtonBusy(reportScope, true);
  try{
    await loadExcelJsForExport();
    const workbook = buildExecutiveExcelMetadataWorkbook(context);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    const suffix = reportScope === 'ejecutivo' ? 'Ejecutivo' : 'Director';
    link.href = url;
    link.download = `Forecast_Ultima_Actualizacion_Excel_${suffix}_${getBogotaTimestampForFile()}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  } catch(error){
    console.error('[EXPORT EXCEL METADATA]', error);
    alert('No se pudo generar el reporte de ultima actualizacion del Excel.');
  } finally {
    setExecutiveExcelMetadataButtonBusy(reportScope, false);
  }
}

function fmtNum(v){return Math.round(v).toLocaleString('es-CO')}
function fmtPct(v){return (v*100).toFixed(1)+'%'}
function abr(v){
  if(v>=1e12) return '$ '+(v/1e12).toFixed(2)+' Bill';
  if(v>=1e9)  return '$ '+(v/1e9).toFixed(2)+' MM';   // miles de millones
  if(v>=1e6)  return '$ '+(v/1e6).toFixed(1)+' M';
  if(v>=1e3)  return '$ '+(v/1e3).toFixed(0)+' K';
  return '$ '+Math.round(v).toLocaleString('es-CO');
}
function escAttr(s){
  return String(s||'')
    .replace(/&/g,'&amp;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;');
}

function jsValueLiteral(value){
  if(value === undefined) return 'undefined';
  return JSON.stringify(value);
}

function jsCall(name){
  const args = Array.from(arguments).slice(1).map(jsValueLiteral).join(',');
  return `${name}(${args})`;
}

function escHtml(s){
  return String(s||'')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function cleanDisplayText(value, fallback='Sin dato'){
  const raw = value === null || value === undefined ? '' : String(value).trim();
  if(!raw) return fallback;
  return raw;
}

function optionHtml(value, label, selected){
  return `<option value="${escAttr(value)}"${selected ? ' selected' : ''}>${escHtml(label)}</option>`;
}

function buildOptionList(items, config){
  const opts = config || {};
  return (items || []).map(item => {
    const value = opts.getValue ? opts.getValue(item) : item;
    const label = opts.getLabel ? opts.getLabel(item) : item;
    const selected = opts.selectedValue !== undefined && String(value) === String(opts.selectedValue);
    return optionHtml(value, label, selected);
  }).join('');
}

function getBogotaDateKey(date){
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(date || new Date()).reduce((acc, part) => {
    if(part.type !== 'literal') acc[part.type] = part.value;
    return acc;
  }, {});
  return `${parts.year}-${parts.month}-${parts.day}`;
}

function getFinanceDefaultPeriod(){
  return getBogotaDateKey().slice(0,7);
}

function getFinanceMonthEnd(period){
  const parts = String(period || getFinanceDefaultPeriod()).split('-').map(Number);
  if(parts.length !== 2 || !parts[0] || !parts[1]) return getBogotaDateKey();
  const lastDay = new Date(parts[0], parts[1], 0).getDate();
  return `${String(parts[0]).padStart(4,'0')}-${String(parts[1]).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
}

function getFinanceDefaultEndDate(period){
  const today = getBogotaDateKey();
  return period === today.slice(0,7) ? today : getFinanceMonthEnd(period);
}

function isFinanceAdmin(){
  const role = CURRENT_USER && CURRENT_USER.role;
  return role === 'gerencia' || role === 'gerencia_director';
}

function getFinanceGoalKey(period){
  return FINANCE_GOAL_KEY_PREFIX + (period || getFinanceDefaultPeriod());
}

function readFinanceGoal(period){
  try {
    const raw = localStorage.getItem(getFinanceGoalKey(period));
    const value = parseFloat(raw);
    return Number.isFinite(value) && value > 0 ? value : 0;
  } catch(_) {
    return 0;
  }
}

function writeFinanceGoal(period, value){
  try {
    localStorage.setItem(getFinanceGoalKey(period), String(Math.max(0, Number(value || 0))));
  } catch(_) {}
}

function ensureFinanceStateDefaults(){
  if(!FINANCE_STATE.period) FINANCE_STATE.period = getFinanceDefaultPeriod();
  if(!FINANCE_STATE.endDate) FINANCE_STATE.endDate = getFinanceDefaultEndDate(FINANCE_STATE.period);
}

function setFinanceInputs(){
  ensureFinanceStateDefaults();
  const periodInput = document.getElementById('finance-period');
  const endInput = document.getElementById('finance-end-date');
  const goalInput = document.getElementById('finance-goal');
  const admin = isFinanceAdmin();
  if(periodInput) periodInput.value = FINANCE_STATE.period;
  if(endInput) {
    endInput.value = FINANCE_STATE.endDate;
    endInput.min = `${FINANCE_STATE.period}-01`;
    endInput.max = getFinanceMonthEnd(FINANCE_STATE.period);
  }
  if(goalInput) {
    goalInput.value = readFinanceGoal(FINANCE_STATE.period) || '';
    goalInput.disabled = !admin;
  }
  document.querySelectorAll('.finance-primary-btn').forEach(btn => {
    btn.disabled = !admin;
    btn.style.opacity = admin ? '' : '.55';
    btn.style.cursor = admin ? '' : 'not-allowed';
  });
}

function setFinancePeriod(value){
  if(!/^\d{4}-\d{2}$/.test(String(value || ''))) return;
  FINANCE_STATE.period = value;
  FINANCE_STATE.endDate = getFinanceDefaultEndDate(value);
  FINANCE_STATE.data = null;
  FINANCE_STATE.error = null;
  setFinanceInputs();
  loadFinanceSummary(true);
}

function setFinanceEndDate(value){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) return;
  FINANCE_STATE.endDate = value;
  FINANCE_STATE.data = null;
  FINANCE_STATE.error = null;
  loadFinanceSummary(true);
}

function saveFinanceGoal(){
  if(!isFinanceAdmin()) {
    alert('Solo gerencia puede editar la meta financiera.');
    return;
  }
  ensureFinanceStateDefaults();
  const input = document.getElementById('finance-goal');
  const value = parseFloat(input && input.value || '0');
  writeFinanceGoal(FINANCE_STATE.period, Number.isFinite(value) ? value : 0);
  renderFinanceSummary();
}

function formatFinancePct(value){
  if(!Number.isFinite(Number(value))) return '—';
  return (Number(value) * 100).toLocaleString('es-CO', {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  }) + '%';
}

function getFinanceCategoryRows(data){
  const byCategory = new Map((data && data.categories || []).map(row => [row.category, row]));
  const ordered = FINANCE_CATEGORY_ORDER.map(category => byCategory.get(category) || {
    category,
    sales: 0,
    cost: 0,
    grossProfit: 0,
    marginPct: null,
    documents: 0
  });
  const extras = (data && data.categories || []).filter(row => !FINANCE_CATEGORY_ORDER.includes(row.category));
  return [...ordered, ...extras];
}

function shouldUseFinanceApi(){
  const host = window.location.hostname;
  return host === 'localhost' || host === '127.0.0.1' || host.endsWith('.vercel.app');
}

function parseFinanceNumber(value){
  if(value === null || value === undefined || value === '') return 0;
  const direct = Number(value);
  if(Number.isFinite(direct)) return direct;
  return parseMonto(value) || 0;
}

function normalizeFinanceCategory(value){
  const raw = cleanDisplayText(value, '');
  const normalized = raw
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .trim();
  if(!normalized) return 'Ventas';
  if(normalized.includes('renta')) return 'Renta de Equipos';
  if(normalized.includes('servicio')) return 'Servicios';
  if(normalized.includes('licenciamiento')) return 'Licenciamiento';
  if(normalized.includes('suministro')) return 'Suministros';
  if(normalized.includes('tecnolog')) return 'Tecnología';
  if(normalized === 'pac' || normalized.includes('pac')) return 'PAC';
  return raw || 'Ventas';
}

function extractFinanceCachePayload(candidate){
  if(candidate && candidate.cacheVersion && candidate.payload) {
    return {
      rows: candidate.payload.data || [],
      meta: candidate.payload.meta || {},
      generatedAt: candidate.generatedAt || candidate.payload.meta && candidate.payload.meta.cacheGeneratedAt || null,
    };
  }
  return {
    rows: candidate && candidate.data || [],
    meta: candidate && candidate.meta || {},
    generatedAt: candidate && (candidate.generatedAt || candidate.meta && candidate.meta.cacheGeneratedAt) || null,
  };
}

function normalizeFinanceCachedRows(rows){
  return (rows || []).map(row => {
    const sign = Number(row.signoDocumento || 1) < 0 ? -1 : 1;
    const totalOriginal = Math.abs(parseFinanceNumber(row.totalOriginal ?? row.total));
    const total = parseFinanceNumber(row.total || totalOriginal * sign);
    const cost = sign >= 0 ? parseFinanceNumber(row.costoProducto) : 0;
    return {
      fechaIso: String(row.fechaIso || '').slice(0,10),
      sign,
      total,
      totalOriginal,
      cost,
      category: normalizeFinanceCategory(row.categoria),
    };
  }).filter(row => row.fechaIso);
}

function summarizeFinanceDocuments(documents, range, sourceInfo){
  const selected = documents.filter(row => row.fechaIso >= range.startDate && row.fechaIso <= range.endDate);
  const positives = selected.filter(row => row.sign >= 0);
  const credits = selected.filter(row => row.sign < 0);
  const categoryMap = new Map();

  positives.forEach(row => {
    const category = row.category || 'Ventas';
    if(!categoryMap.has(category)) {
      categoryMap.set(category, { category, sales: 0, cost: 0, grossProfit: 0, marginPct: 0, documents: 0 });
    }
    const bucket = categoryMap.get(category);
    bucket.sales += row.totalOriginal;
    bucket.cost += row.cost;
    bucket.documents += 1;
  });

  const categories = [...categoryMap.values()].map(row => ({
    ...row,
    grossProfit: row.sales - row.cost,
    marginPct: row.sales ? (row.sales - row.cost) / row.sales : 0,
  })).sort((a,b) => b.sales - a.sales);

  const grossSales = positives.reduce((sum, row) => sum + row.totalOriginal, 0);
  const creditNotes = credits.reduce((sum, row) => sum + row.totalOriginal, 0);
  const netSales = selected.reduce((sum, row) => sum + row.total, 0);
  const cost = positives.reduce((sum, row) => sum + row.cost, 0);
  const grossProfit = grossSales - cost;

  return {
    ok: true,
    period: range.period,
    startDate: range.startDate,
    endDate: range.endDate,
    source: sourceInfo.source || 'cache-browser',
    sourceName: sourceInfo.sourceName || 'Cache PBI ventas',
    sourceGeneratedAt: sourceInfo.generatedAt || null,
    sourceRange: sourceInfo.range || null,
    totalAvailableDocuments: documents.length,
    rows: selected.length,
    positiveDocuments: positives.length,
    creditDocuments: credits.length,
    totals: {
      grossSales,
      creditNotes,
      netSales,
      cost,
      grossProfit,
      marginPct: grossSales ? grossProfit / grossSales : 0,
    },
    categories,
  };
}

async function requestFinanceSummaryFromApi(params){
  const response = await fetch(`/api/finance-summary?${params.toString()}`, { cache: 'no-store' });
  const payload = await response.json().catch(() => ({}));
  if(!response.ok || !payload.ok) {
    throw new Error(payload.error || `HTTP ${response.status}`);
  }
  return payload;
}

async function requestFinanceSummaryFromCache(range){
  const response = await fetch(FINANCE_PUBLIC_CACHE_URL, { cache: 'no-store' });
  if(!response.ok) throw new Error(`Cache PBI HTTP ${response.status}`);
  const candidate = await response.json();
  const payload = extractFinanceCachePayload(candidate);
  const documents = normalizeFinanceCachedRows(payload.rows);
  return summarizeFinanceDocuments(documents, range, {
    source: 'cache-browser',
    sourceName: payload.meta && payload.meta.sourceName || 'Cache PBI ventas',
    generatedAt: payload.generatedAt,
    range: payload.meta && payload.meta.range || null,
  });
}

async function loadFinanceSummary(force){
  ensureFinanceStateDefaults();
  if(FINANCE_STATE.isLoading) return;
  if(FINANCE_STATE.data && !force) {
    renderFinanceSummary();
    return;
  }
  FINANCE_STATE.isLoading = true;
  FINANCE_STATE.error = null;
  renderFinanceSummary();
  try {
    const params = new URLSearchParams({
      period: FINANCE_STATE.period,
      startDate: `${FINANCE_STATE.period}-01`,
      endDate: FINANCE_STATE.endDate
    });
    const range = {
      period: FINANCE_STATE.period,
      startDate: `${FINANCE_STATE.period}-01`,
      endDate: FINANCE_STATE.endDate
    };
    let payload = null;
    if(shouldUseFinanceApi()) {
      try {
        payload = await requestFinanceSummaryFromApi(params);
      } catch(apiError) {
        console.warn('[FINANCE] API local no disponible, usando cache publico', apiError);
      }
    }
    if(!payload) payload = await requestFinanceSummaryFromCache(range);
    FINANCE_STATE.data = payload;
  } catch(error) {
    FINANCE_STATE.error = error instanceof Error ? error.message : 'No fue posible cargar el resumen financiero.';
    FINANCE_STATE.data = null;
  } finally {
    FINANCE_STATE.isLoading = false;
    renderFinanceSummary();
  }
}

function renderFinanceAdminNote(){
  const note = document.getElementById('finance-admin-note');
  if(!note) return;
  if(isFinanceAdmin()) {
    note.textContent = 'Meta editable para gerencia. En esta primera versión la meta queda guardada por mes en este navegador.';
  } else {
    note.textContent = 'Vista de consulta. La edición de meta está reservada para gerencia.';
  }
}

function renderFinanceSourceNote(data){
  const note = document.getElementById('finance-source-note');
  if(!note) return;
  if(!data) {
    note.textContent = '';
    return;
  }
  const generated = data.sourceGeneratedAt ? formatLastConnectionFull(data.sourceGeneratedAt) : 'Sin fecha de cache';
  const range = data.sourceRange && data.sourceRange.start && data.sourceRange.end
    ? `Rango fuente: ${data.sourceRange.start} a ${data.sourceRange.end}.`
    : '';
  note.textContent = `Fuente: ${data.sourceName || 'PBI ventas'} (${data.source || 'api'}). Corte consultado: ${data.startDate} a ${data.endDate}. Cache/fuente: ${generated}. ${range}`;
}

function renderFinanceSummary(){
  ensureFinanceStateDefaults();
  setFinanceInputs();
  renderFinanceAdminNote();
  const kpis = document.getElementById('finance-kpis');
  const formulaProfit = document.getElementById('finance-profit-formula');
  const formulaMargin = document.getElementById('finance-margin-formula');
  const table = document.getElementById('finance-category-table');
  if(!kpis || !formulaProfit || !formulaMargin || !table) return;

  if(FINANCE_STATE.isLoading) {
    kpis.innerHTML = `<div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div><div class="kpi-label">Cargando</div><div class="kpi-val">PBI</div><div class="kpi-sub">Consultando resumen financiero</div></div>`;
    formulaProfit.textContent = 'Cargando datos PBI...';
    formulaMargin.textContent = 'Cargando datos PBI...';
    table.innerHTML = '<div style="padding:18px;color:var(--text2)">Consultando datos de ventas.</div>';
    renderFinanceSourceNote(null);
    return;
  }

  if(FINANCE_STATE.error) {
    kpis.innerHTML = `<div class="kpi" style="--ac:var(--corp-red)"><div class="kpi-accent"></div><div class="kpi-label">Error</div><div class="kpi-val">Sin datos</div><div class="kpi-sub">${escHtml(FINANCE_STATE.error)}</div></div>`;
    formulaProfit.textContent = 'No fue posible calcular la utilidad.';
    formulaMargin.textContent = 'No fue posible calcular el margen.';
    table.innerHTML = `<div style="padding:18px;color:var(--text2)">${escHtml(FINANCE_STATE.error)}</div>`;
    renderFinanceSourceNote(null);
    return;
  }

  const data = FINANCE_STATE.data;
  if(!data) {
    kpis.innerHTML = `<div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div><div class="kpi-label">Finanzas</div><div class="kpi-val">Listo</div><div class="kpi-sub">Define la meta y actualiza el resumen.</div></div>`;
    formulaProfit.textContent = 'Ventas acumuladas - Costo PBI = Utilidad bruta';
    formulaMargin.textContent = 'Utilidad bruta / Ventas acumuladas = Margen bruto';
    table.innerHTML = '<div style="padding:18px;color:var(--text2)">Aún no hay datos cargados.</div>';
    renderFinanceSourceNote(null);
    return;
  }

  const goal = readFinanceGoal(FINANCE_STATE.period);
  const totals = data.totals || {};
  const sales = Number(totals.grossSales || 0);
  const cost = Number(totals.cost || 0);
  const grossProfit = Number(totals.grossProfit || 0);
  const marginPct = Number(totals.marginPct || 0);
  const fulfillment = goal > 0 ? sales / goal : null;

  kpis.innerHTML = `
    <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
      <div class="kpi-label">Proyección ventas ${escHtml(FINANCE_STATE.period)}</div>
      <div class="kpi-val">${goal ? abr(goal) : '—'}</div>
      <div class="kpi-sub">${goal ? fmtCOP(goal) : 'Meta pendiente'}</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
      <div class="kpi-label">Ventas 01 al ${escHtml(FINANCE_STATE.endDate.slice(8,10))}</div>
      <div class="kpi-val">${abr(sales)}</div>
      <div class="kpi-sub">${fmtCOP(sales)} · ${fmtNum(data.positiveDocuments || 0)} FV</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
      <div class="kpi-label">Cumplimiento</div>
      <div class="kpi-val">${fulfillment === null ? '—' : formatFinancePct(fulfillment)}</div>
      <div class="kpi-sub">Ventas acumuladas / meta</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
      <div class="kpi-label">Utilidad bruta total</div>
      <div class="kpi-val">${abr(grossProfit)}</div>
      <div class="kpi-sub">${fmtCOP(grossProfit)}</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
      <div class="kpi-label">Margen bruto del mes</div>
      <div class="kpi-val">${formatFinancePct(marginPct)}</div>
      <div class="kpi-sub">Utilidad bruta / ventas</div>
    </div>
  `;

  formulaProfit.textContent = `${fmtCOP(sales)} - ${fmtCOP(cost)} = ${fmtCOP(grossProfit)}`;
  formulaMargin.textContent = `${fmtCOP(grossProfit)} / ${fmtCOP(sales)} = ${formatFinancePct(marginPct)}`;

  const rows = getFinanceCategoryRows(data);
  table.innerHTML = `<table class="responsive-table">
    <thead>
      <tr><th>Categoría</th><th>Documentos</th><th>Ventas</th><th>Costo PBI</th><th>Utilidad</th><th>% Utilidad</th></tr>
    </thead>
    <tbody>${rows.map((row, index) => {
      const color = FINANCE_CATEGORY_COLORS[index % FINANCE_CATEGORY_COLORS.length];
      const rowSales = Number(row.sales || 0);
      const rowCost = Number(row.cost || 0);
      const rowProfit = Number(row.grossProfit || rowSales - rowCost || 0);
      const rowMargin = rowSales ? rowProfit / rowSales : null;
      return `<tr>
        <td data-label="Categoría">
          <div class="finance-category-cell">
            <span class="finance-category-swatch" style="background:${escAttr(color)}"></span>
            <span>${escHtml(row.category)}</span>
          </div>
        </td>
        <td class="td-mono" data-label="Documentos">${fmtNum(row.documents || 0)}</td>
        <td class="td-mono td-cop" data-label="Ventas">${rowSales ? fmtCOP(rowSales) : '—'}</td>
        <td class="td-mono" data-label="Costo PBI">${rowCost ? fmtCOP(rowCost) : '—'}</td>
        <td class="td-mono" data-label="Utilidad">${rowSales ? fmtCOP(rowProfit) : '—'}</td>
        <td class="td-mono" data-label="% Utilidad">${rowMargin === null ? '—' : formatFinancePct(rowMargin)}</td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;

  renderFinanceSourceNote(data);
}

function renderFinanzas(){
  ensureFinanceStateDefaults();
  renderFinanceSummary();
  if(!FINANCE_STATE.data && !FINANCE_STATE.error && !FINANCE_STATE.isLoading) {
    loadFinanceSummary(false);
  }
}

function normalizeHeaderKey(v){
  return String(v || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toUpperCase()
    .replace(/\s+/g,' ')
    .trim();
}

const HEADER_KEY_MAP = {
  'CLIENTE':'CLIENTE',
  'COMERCIAL':'COMERCIAL',
  'SALES SUPPORT':'SALES SUPPORT',
  'SALESSUPPORT':'SALES SUPPORT',
  'VENTA CLIENTE':'MONTO VENTA CLIENTE',
  'MONTO VENTA':'MONTO VENTA CLIENTE',
  'VALOR VENTA':'MONTO VENTA CLIENTE',
  'VALOR VENTA CLIENTE':'MONTO VENTA CLIENTE',
  'VALOR CLIENTE':'MONTO VENTA CLIENTE',
  'NUMERO DE COTIZACION':'NUMERO DE COTIZACION',
  'NUMERO COTIZACION':'NUMERO DE COTIZACION',
  'NRO COTIZACION':'NUMERO DE COTIZACION',
  'NO COTIZACION':'NUMERO DE COTIZACION',
  'COTIZACION':'NUMERO DE COTIZACION',
  'COSTO NEGOCIO':'COSTO NEGOCIO',
  'NUMERO PARTE':'NUMERO DE PARTE',
  'NRO PARTE':'NUMERO DE PARTE',
  'NO PARTE':'NUMERO DE PARTE',
  'PART NUMBER':'NUMERO DE PARTE',
  'CANT':'CANTIDAD',
  'CANT.':'CANTIDAD',
  'QTY':'CANTIDAD',
  'CANTIDAD':'CANTIDAD',
  'CANTIDADES':'CANTIDAD',
  'UNIDADES':'CANTIDAD',
  'UND':'CANTIDAD',
  'UNIDAD':'CANTIDAD',
  'CANTIDAD PRODUCTO':'CANTIDAD',
  'CANTIDAD SOLICITADA':'CANTIDAD',
  'COSTO DEL NEGOCIO':'COSTO NEGOCIO',
  'COSTO TOTAL NEGOCIO':'COSTO NEGOCIO',
  'MONEDA':'MONEDA 2',
  'DIVISA':'MONEDA 2',
  'MONEDA2':'MONEDA 2',
  'ESTADO':'ESTADO',
  'UTILIDAD':'UTILIDAD',
  'MARGEN':'MARGEN',
  'SOPORTA':'SOPORTA',
  'SOPORTADO':'SOPORTA',
  'COMERCIAL APOYADO':'SOPORTA',
  'FECHA':'FECHA DIA/MES/AÑO',
  'FECHA NEGOCIO':'FECHA DIA/MES/AÑO',
  'FECHA VENTA':'FECHA DIA/MES/AÑO',
  'TRM':'TRM REFERENCIA',
  'TRM REF':'TRM REFERENCIA',
  'TRM DE REFERENCIA':'TRM REFERENCIA',
  'LINEA':'LINEA DE PRODUCTO',
  'LÍNEA':'LINEA DE PRODUCTO',
  'LÍNEA DE PRODUCTO':'LINEA DE PRODUCTO',
  'COSTO TOTAL':'COSTO',
  'VENDEDOR':'COMERCIAL',
  'EJECUTIVO':'COMERCIAL',
  'MODIFICACION':'MODIFICACION O UPGRADE',
  'MODIFICACIÓN':'MODIFICACION O UPGRADE',
  'OBSERVACION':'OBSERVACIONES',
  'OBSERVACIONES':'OBSERVACIONES',
  'OBSERVACION COMERCIAL':'OBSERVACIONES',
  'OBSERVACIONES COMERCIALES':'OBSERVACIONES',
  'COMENTARIO':'OBSERVACIONES',
  'COMENTARIOS':'OBSERVACIONES',
  'NOTAS':'OBSERVACIONES',
  'PENDIENTE':'PENDIENTE',
  'PENDIENTES':'PENDIENTE',
  'TIPO PENDIENTE':'PENDIENTE',
  'TIPO DE PENDIENTE':'PENDIENTE',
  'NUMERO PENDIENTE':'NUMERO PENDIENTE',
  'NUMERO DE PENDIENTE':'NUMERO PENDIENTE',
  'NRO PENDIENTE':'NUMERO PENDIENTE',
  'NRO. PENDIENTE':'NUMERO PENDIENTE',
  'NO PENDIENTE':'NUMERO PENDIENTE',
  'NO. PENDIENTE':'NUMERO PENDIENTE',
  'NUM PENDIENTE':'NUMERO PENDIENTE',
  'N PENDIENTE':'NUMERO PENDIENTE',
  'N° PENDIENTE':'NUMERO PENDIENTE',
  'ID PENDIENTE':'NUMERO PENDIENTE',
  '# PENDIENTE':'NUMERO PENDIENTE',
  'CATEGORIA':'CATEGORIA',
  'CATEGORÍA':'CATEGORIA',
  'VALOR FACTURAS':'VALOR FACTURAS',
  'VALOR FACTURA':'VALOR FACTURAS',
  'VALOR DE FACTURAS':'VALOR FACTURAS',
  'VALOR DE FACTURA':'VALOR FACTURAS',
  'VALOR FACTURADO':'VALOR FACTURAS',
  'VALOR FACTURACION':'VALOR FACTURAS',
  'VALOR FACTURACIÓN':'VALOR FACTURAS',
  'VALOR FACTURAS COP':'VALOR FACTURAS',
  'VALOR FACTURA COP':'VALOR FACTURAS',
  'VALOR FACTURAS $':'VALOR FACTURAS',
  'FACTURA':'VALOR FACTURAS'
};

const DETAIL_FIELD_LABELS = {
  CLIENTE: 'Cliente',
  'NOMBRE CLIENTE': 'Nombre cliente',
  EMPRESA: 'Empresa',
  PRODUCTO: 'Proyecto o producto',
  PROYECTO: 'Proyecto o producto',
  SOLUCION: 'Solucion',
  SERVICIO: 'Servicio',
  'NUMERO DE PARTE': 'Numero de parte',
  CANTIDAD: 'Cantidad',
  'NUMERO DE COTIZACION': 'Numero de cotizacion',
  MARCA: 'Marca',
  FABRICANTE: 'Marca',
  'LINEA DE PRODUCTO': 'Linea',
  ESTADO: 'Estado',
  DIRECTOR: 'Director',
  COMERCIAL: 'Ejecutivo',
  'SALES SUPPORT': 'Sales Support',
  SOPORTA: 'Soporta',
  'MONEDA 2': 'Moneda',
  'FECHA DIA/MES/AÑO': 'Fecha',
  'MONTO VENTA CLIENTE': 'Valor negocio',
  COSTO: 'Costo',
  'COSTO NEGOCIO': 'Costo negocio',
  UTILIDAD: 'Utilidad',
  'TRM REFERENCIA': 'TRM del dia',
  MARGEN: 'Margen',
  'MODIFICACION O UPGRADE': 'Modificacion o upgrade',
  OBSERVACIONES: 'Observaciones',
  PENDIENTE: 'Pendiente',
  'NUMERO PENDIENTE': 'Numero pendiente',
  CATEGORIA: 'Categoria',
  'VALOR FACTURAS': 'Valor facturas'
};

function mapHeaderName(h){
  const raw = String(h || '').trim();
  const normalized = normalizeHeaderKey(raw);
  return HEADER_KEY_MAP[normalized] || raw;
}

function registerRecord(rec, sourceFile, sourceSheet, dataset){
  const ds = dataset === 'sales'
    ? 'sales'
    : dataset === 'sales_pending'
      ? 'sales_pending'
      : dataset === 'preventa'
        ? 'preventa'
        : 'forecast';
  rec.__RID = (ds === 'sales'
    ? 'sales-'
    : ds === 'sales_pending'
      ? 'sales-pending-'
      : ds === 'preventa'
        ? 'preventa-'
        : 'neg-') + (++RECORD_SEQ);
  rec.__DATASET = ds;
  rec.__SOURCE_FILE = sourceFile || '';
  rec.__SOURCE_SHEET = sourceSheet || '';
  return rec;
}

function firstFilled(row, keys){
  for(const key of keys){
    const val = row && row[key];
    if(val !== null && val !== undefined && String(val).trim() !== '') return val;
  }
  return '';
}

function getRowClientName(row){
  return firstFilled(row, ['CLIENTE','NOMBRE CLIENTE','EMPRESA','RAZON SOCIAL']) || '';
}

function getRowProductName(row){
  return firstFilled(row, ['PRODUCTO','PROYECTO','SOLUCION','SERVICIO']) || '';
}

const BRAND_NAME_ALIASES = {
  acer: 'ACER',
  adobe: 'ADOBE',
  apple: 'APPLE',
  aruba: 'ARUBA',
  asus: 'ASUS',
  bitdefender: 'BITDEFENDER',
  cisco: 'CISCO',
  dahua: 'DAHUA',
  dell: 'DELL',
  epson: 'EPSON',
  fortinet: 'FORTINET',
  fortine: 'FORTINET',
  generica: 'GENERICA',
  generico: 'GENERICA',
  hikvision: 'HIKVISION',
  honeywell: 'HONEYWELL',
  hp: 'HP',
  'hp inc': 'HP',
  'hp inc.': 'HP',
  hpe: 'HP',
  'hewlett-packard': 'HP',
  'hewlett packard': 'HP',
  jabra: 'JABRA',
  kingston: 'KINGSTON',
  kinstong: 'KINGSTON',
  kiinstong: 'KINGSTON',
  kyocera: 'KYOCERA',
  kiocera: 'KYOCERA',
  lenovo: 'LENOVO',
  lg: 'LG',
  logitech: 'LOGITECH',
  loguitech: 'LOGITECH',
  logutech: 'LOGITECH',
  microsof: 'MICROSOFT',
  microsoft: 'MICROSOFT',
  motorola: 'MOTOROLA',
  multimarca: 'MULTIMARCA',
  multimarcas: 'MULTIMARCA',
  poly: 'POLY',
  qnap: 'QNAP',
  ricoh: 'RICOH',
  samsung: 'SAMSUNG',
  samsumg: 'SAMSUNG',
  samsusng: 'SAMSUNG',
  sony: 'SONY',
  '3n star': '3N STAR',
  '3nstar': '3N STAR',
  'tp-link': 'TP-LINK',
  tplink: 'TP-LINK',
  ubiquiti: 'UBIQUITI',
  ubituiti: 'UBIQUITI',
  varios: 'VARIOS',
  varias: 'VARIOS',
  viewsonic: 'VIEWSONIC',
  xiaomi: 'XIAOMI',
  zebra: 'ZEBRA'
};

function isCompositeBrandName(value){
  return /,|;|\+|\/|\sy\s|(?:\s-\s*|\s*-\s)/i.test(value || '');
}

function normalizeBrandDisplay(value){
  return String(value || '')
    .replace(/\s*,\s*/g, ', ')
    .replace(/\s*;\s*/g, '; ')
    .replace(/\s*\/\s*/g, ' / ')
    .replace(/\s*\+\s*/g, ' + ')
    .replace(/\s+y\s+/ig, ' Y ')
    .replace(/\s*-\s*/g, ' - ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function normalizeBrandName(value){
  const raw = cleanDisplayText(value, '');
  if(!raw) return '';
  const compact = raw.replace(/\s+/g, ' ').trim();
  const aliasKey = normalizeCategoryValue(compact);
  if(BRAND_NAME_ALIASES[aliasKey]) return BRAND_NAME_ALIASES[aliasKey];
  if(isCompositeBrandName(compact)) return normalizeBrandDisplay(compact);
  return compact.toUpperCase();
}

function getRowBrandName(row){
  return normalizeBrandName(firstFilled(row, ['MARCA','FABRICANTE']) || '');
}

function getRowPartNumber(row){
  return firstFilled(row, ['NUMERO DE PARTE','NUMERO PARTE','NRO PARTE','NO PARTE','PART NUMBER']) || '';
}

function getRowQuantityValue(row, fallback){
  const defaultValue = fallback === undefined ? 1 : fallback;
  const raw = firstFilled(row, [
    'CANTIDAD','CANT','CANT.','QTY','UNIDADES','UND','UNIDAD',
    'CANTIDAD PRODUCTO','CANTIDAD SOLICITADA'
  ]);
  if(raw === null || raw === undefined || String(raw).trim() === '') return defaultValue;
  const parsed = parseMonto(raw);
  return parsed > 0 ? parsed : defaultValue;
}

const LINE_NAME_ALIASES = {
  tecnologia: 'TECNOLOGIA',
  tecnolgia: 'TECNOLOGIA',
  tcnologia: 'TECNOLOGIA'
};

function normalizeLineName(value){
  const raw = cleanDisplayText(value, '');
  if(!raw) return '';
  const display = raw
    .replace(/\s*\/\s*/g, ' / ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
  const aliasKey = normalizeCategoryValue(display);
  return LINE_NAME_ALIASES[aliasKey] || display;
}

function getRowLineName(row){
  return normalizeLineName(firstFilled(row, ['LINEA DE PRODUCTO','LINEA']) || '');
}

function getRowObservation(row){
  return firstFilled(row, ['OBSERVACIONES','OBSERVACION','COMENTARIOS','COMENTARIO']);
}

function getSalesSupportName(row){
  return canonicalizeSalesSupportName(firstFilled(row, ['SALES SUPPORT','COMERCIAL']) || '');
}

function splitSalesTargets(value){
  return cleanDisplayText(value, '')
    .split(/\s+-\s+/)
    .map(cleanNameSegment)
    .filter(Boolean);
}

function formatSalesTargets(value){
  const names = Array.isArray(value) ? value : splitSalesTargets(value);
  return names.join(', ');
}

function getSalesSoportaNames(row){
  const raw = firstFilled(row, ['SOPORTA','SOPORTADO','COMERCIAL APOYADO','DIRECTOR APOYADO']) || '';
  return splitSalesTargets(raw);
}

function getSalesSoportaName(row){
  return formatSalesTargets(getSalesSoportaNames(row));
}

function getSalesQuoteNumber(row){
  return firstFilled(row, ['NUMERO DE COTIZACION','NUMERO COTIZACION','NRO COTIZACION','COTIZACION']) || '';
}

function getSalesPendingSupportName(row){
  return canonicalizeSalesSupportName(firstFilled(row, ['SALES SUPPORT']) || '');
}

function getSalesPendingType(row){
  return cleanDisplayText(firstFilled(row, ['PENDIENTE']), 'Sin pendiente');
}

function getSalesPendingNumber(row){
  return firstFilled(row, ['NUMERO PENDIENTE']) || '';
}

function getSalesPendingCategory(row){
  return cleanDisplayText(firstFilled(row, ['CATEGORIA']), 'Sin categoria');
}

function getSalesPendingCommercial(row){
  return cleanNameSegment(firstFilled(row, ['COMERCIAL','SOPORTA']) || '');
}

function getSalesPendingInvoiceValue(row){
  return parseMonto(firstFilled(row, ['VALOR FACTURAS'])) || 0;
}

function ensureSalesPendingSupportName(rows, supportName){
  const name = canonicalizeSalesSupportName(supportName || '');
  if(!name) return rows || [];
  (rows || []).forEach(row => {
    if(!getSalesPendingSupportName(row)) row['SALES SUPPORT'] = name;
  });
  return rows || [];
}

function getSalesCostValue(row){
  return parseMonto(firstFilled(row, ['COSTO NEGOCIO','COSTO'])) || 0;
}

function getSalesUtilidadValue(row){
  const raw = firstFilled(row, ['UTILIDAD']);
  if(raw !== '' && raw !== null && raw !== undefined) {
    return parseMonto(raw) || 0;
  }
  return getUtilidad(row).valor;
}

function getRowDateValue(row){
  return firstFilled(row, ['FECHA DIA/MES/AÑO','FECHA']);
}

function formatDateValue(v){
  if(!v) return '';
  if(v instanceof Date) return v.toISOString().substring(0,10);
  if(typeof v === 'number') {
    return new Date(Math.round((v - 25569) * 86400 * 1000)).toISOString().substring(0,10);
  }
  const s = String(v).trim();
  const meses = {ene:'01',feb:'02',mar:'03',abr:'04',may:'05',jun:'06',jul:'07',ago:'08',sep:'09',oct:'10',nov:'11',dic:'12'};
  const m1 = s.match(/(\d{2})[-/](\w{3})\.?[-/](\d{2,4})/i);
  if(m1){
    const mes = meses[m1[2].toLowerCase()] || '01';
    const anio = m1[3].length === 2 ? '20' + m1[3] : m1[3];
    return anio + '-' + mes + '-' + m1[1].padStart(2,'0');
  }
  return s.substring(0,10);
}

function getRecordById(id){
  return ALL_DATA.find(r => r.__RID === id)
    || SALES_DATA.find(r => r.__RID === id)
    || SALES_PENDING_DATA.find(r => r.__RID === id)
    || PREVENTA_DATA.find(r => r.__RID === id)
    || null;
}

function formatFieldLabel(key){
  const mapped = mapHeaderName(key);
  const normalized = normalizeHeaderKey(mapped);
  return DETAIL_FIELD_LABELS[normalized] || cleanDisplayText(mapped, cleanDisplayText(key, 'Campo'));
}

function normalizePersonName(v){
  return (v||'')
    .toString()
    .replace(/\.(xlsx|xls)$/i,'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}

function namesMatch(a,b){
  const na = normalizePersonName(a);
  const nb = normalizePersonName(b);
  if(!na || !nb) return false;
  if(na === nb) return true;

  // Permite casos como "dayana chala" vs "dayana chala perez"
  if((na.includes(nb) || nb.includes(na)) && Math.min(na.length, nb.length) >= 8) {
    return true;
  }

  const ta = [...new Set(na.split(' ').filter(Boolean))];
  const tb = [...new Set(nb.split(' ').filter(Boolean))];
  const common = ta.filter(t => tb.includes(t));

  if(common.length >= 2) return true;
  if(common.length === 1 && (ta.length === 1 || tb.length === 1) && common[0].length >= 5) return true;

  return false;
}

function getSalesSupportAliasMap(){
  return window.SALES_SUPPORT_NAME_ALIASES || {};
}

function canonicalizeSalesSupportName(name){
  const cleanName = cleanNameSegment(name);
  if(!cleanName) return '';
  const aliasMap = getSalesSupportAliasMap();
  for(const [canonicalName, aliases] of Object.entries(aliasMap)) {
    if(namesMatch(cleanName, canonicalName)) return canonicalName;
    if((aliases || []).some(alias => namesMatch(cleanName, alias))) return canonicalName;
  }
  return cleanName;
}

function getSalesSupportTargetNames(){
  const targetName = getSalesSupportTargetName();
  const aliasMap = getSalesSupportAliasMap();
  const aliases = aliasMap[targetName] || [];
  return [...new Set([targetName, ...aliases].map(cleanNameSegment).filter(Boolean))];
}

function matchesAnySalesSupportName(name, candidates){
  const list = Array.isArray(candidates) ? candidates : [];
  return list.some(candidate => namesMatch(name, candidate));
}

/* Get TRM */
function getTRM(){
  const raw = document.getElementById('trm-input').value;
  const v = parseFloat(String(raw).replace(',', '.'));
  return v || 4150;
}

/* Liquidate to COP */
function parseMonto(v){
  if(v===null||v===undefined) return 0;
  if(typeof v==='number') return v;
  let s = String(v).trim().replace(/[$\s]/g,'');
  if(!s) return 0;
  // Detectar formato: colombiano "98.544.430" o americano "98,544.43" o decimal "33262.55"
  const dots = (s.match(/\./g)||[]).length;
  const commas = (s.match(/,/g)||[]).length;
  if(dots > 1){
    // Múltiples puntos = separadores de miles colombianos: "98.544.430" → 98544430
    s = s.replace(/\./g,'').replace(',','.');
  } else if(commas > 1){
    // Múltiples comas = separadores de miles americanos: "98,544,430" → 98544430
    s = s.replace(/,/g,'');
  } else if(dots === 1 && commas === 1){
    // Ambos: puede ser "1.234,56" (EU) o "1,234.56" (US)
    const dotPos = s.indexOf('.');
    const commaPos = s.indexOf(',');
    if(commaPos < dotPos){ s = s.replace(/,/g,''); } // US: 1,234.56
    else { s = s.replace(/\./g,'').replace(',','.'); } // EU: 1.234,56
  } else {
    // Un solo punto o coma — es separador decimal
    s = s.replace(',','.');
  }
  return parseFloat(s)||0;
}

function parsefecha(v){
  if(!v) return '';
  if(v instanceof Date) return v.toISOString().substring(0,7);
  if(typeof v==='number') {
    const d = new Date(Math.round((v - 25569)*86400*1000));
    return d.toISOString().substring(0,7);
  }
  const s = String(v);
  const meses = {ene:'01',feb:'02',mar:'03',abr:'04',may:'05',jun:'06',jul:'07',ago:'08',sep:'09',oct:'10',nov:'11',dic:'12'};
  const m1 = s.match(/(\d{2})[-/](\w{3})\.?[-/](\d{2,4})/i);
  if(m1){
    const mes = meses[m1[2].toLowerCase()]||'01';
    const anio = m1[3].length===2?'20'+m1[3]:m1[3];
    return anio+'-'+mes;
  }
  return s.substring(0,7);
}

function normalizeEstado(v){
  if(v===null||v===undefined) return '';
  const s = String(v).trim().toUpperCase();
  if(s === 'GANADO') return 'GANADA';
  if(s === 'PERDIDO' || s === 'PEDIDA') return 'PERDIDA';
  if(s === 'APLAZADA') return 'APLAZADO';
  return s;
}

function getEstadoBadgeClass(value){
  const estado = normalizeEstado(value);
  return ['GANADA','PENDIENTE','PERDIDA','APLAZADO'].includes(estado) ? estado : 'PENDIENTE';
}

let _chartTooltip = null;
function getChartTooltip(){
  if(_chartTooltip) return _chartTooltip;
  const tip = document.createElement('div');
  tip.className = 'chart-tooltip';
  tip.setAttribute('role','tooltip');
  document.body.appendChild(tip);
  _chartTooltip = tip;
  return tip;
}

function positionChartTooltip(e, tip){
  const pad = 10;
  let x = e.clientX + 14;
  let y = e.clientY - 12;
  tip.style.left = x + 'px';
  tip.style.top = y + 'px';
  const r = tip.getBoundingClientRect();
  if(r.right > window.innerWidth - pad) {
    x = window.innerWidth - r.width - pad;
  }
  if(r.top < pad) {
    y = e.clientY + 18;
  }
  tip.style.left = x + 'px';
  tip.style.top = y + 'px';
}

function attachChartTooltips(container){
  if(!container) return;
  const nodes = container.querySelectorAll('[data-tooltip]');
  if(!nodes.length) return;
  const tip = getChartTooltip();
  nodes.forEach(node => {
    if(node.dataset.ttBound) return;
    node.dataset.ttBound = '1';
    node.addEventListener('pointerenter', e => {
      const text = node.getAttribute('data-tooltip');
      if(!text) return;
      tip.textContent = text;
      tip.classList.add('show');
      positionChartTooltip(e, tip);
    });
    node.addEventListener('pointermove', e => {
      if(tip.classList.contains('show')) positionChartTooltip(e, tip);
    });
    node.addEventListener('pointerleave', () => {
      tip.classList.remove('show');
    });
  });
}

function toCOP(row){
  const m = (row['MONEDA 2']||'COP').trim().toUpperCase();
  const val = parseMonto(row['MONTO VENTA CLIENTE']);
  const trm = getTRM();
  if(m==='USD') return val * trm;
  return val;
}

function getMarginRatio(raw){
  if(raw === null || raw === undefined || String(raw).trim() === '') return null;
  const parsed = parseMonto(raw);
  if(isNaN(parsed)) return null;
  if(parsed >= 0 && parsed <= 1) return parsed;
  if(parsed > 1 && parsed <= 100) return parsed / 100;
  return null;
}

function formatMarginDisplay(raw, fallback='Sin dato'){
  const ratio = getMarginRatio(raw);
  return ratio === null ? fallback : fmtPct(ratio);
}

function getUtilidad(row){
  const moneda = cleanDisplayText(row['MONEDA 2'], 'COP').trim().toUpperCase();
  const utilidadRaw = firstFilled(row, ['UTILIDAD']);
  if(utilidadRaw !== '' && utilidadRaw !== null && utilidadRaw !== undefined) {
    return { valor: parseMonto(utilidadRaw) || 0, moneda };
  }
  const valor = parseMonto(row['MONTO VENTA CLIENTE']);
  const margen = getMarginRatio(row['MARGEN']);
  if(!(valor > 0) || margen === null) return { valor: 0, moneda };
  return { valor: valor * margen, moneda };
}

function toTitleName(s){
  if(!s) return '';
  return String(s).replace(/^[' ]+/,'').trim()
    .replace(/\w\S*/g,w=>w.charAt(0).toUpperCase()+w.slice(1).toLowerCase());
}

function cleanNameSegment(value){
  return cleanDisplayText(value, '').replace(/\s+/g,' ').trim();
}

function normalizeDirectorName(value){
  const normalized = normalizePersonName(value);
  if(!normalized) return '';
  return normalized
    .split(' ')
    .filter(Boolean)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

function normalizeFileBaseName(name){
  return String(name || '')
    .replace(/\.(xlsx|xls)$/i,'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .replace(/[_\-–—.]+/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}

function isSalesSupportFile(name){
  const base = normalizeFileBaseName(name);
  return /^sales\s*support(?:\s|$)/i.test(base);
}

function normalizeSheetName(name){
  return String(name || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .replace(/\s+/g,'')
    .trim();
}

function pickWorksheetName(sheetNames, datasetType){
  const names = Array.isArray(sheetNames) ? sheetNames : [];
  if(!names.length) return '';
  const preferred = datasetType === 'sales'
    ? ['salesreport','sales support','salessupport','sales report','comercial','gerencia']
    : ['gerencia','comercial'];

  for(const needle of preferred){
    const normNeedle = normalizeSheetName(needle);
    const match = names.find(name => normalizeSheetName(name).includes(normNeedle));
    if(match) return match;
  }
  return names[0];
}

function pickSalesPendingWorksheetName(sheetNames){
  const names = Array.isArray(sheetNames) ? sheetNames : [];
  if(!names.length) return '';
  const exact = names.find(name => normalizeSheetName(name) === 'pendientes');
  if(exact) return exact;
  return names.find(name => normalizeSheetName(name).includes('pendiente')) || '';
}

function isLikelyHeaderRow(row){
  if(!Array.isArray(row) || !row.length) return false;
  const headers = row.map(cell => normalizeHeaderKey(cell)).filter(Boolean);
  if(!headers.length) return false;
  const hasCliente = headers.some(h => h.includes('CLIENTE'));
  const hasEstado = headers.includes('ESTADO');
  const hasMoneda = headers.includes('MONEDA') || headers.includes('MONEDA 2');
  const hasValor = headers.some(h => h.includes('VENTA CLIENTE') || h.includes('MONTO VENTA') || h.includes('VALOR VENTA'));
  return hasCliente && (hasEstado || hasMoneda || hasValor);
}

function isLikelySalesPendingHeaderRow(row){
  if(!Array.isArray(row) || !row.length) return false;
  const headers = row.map(cell => normalizeHeaderKey(mapHeaderName(cell))).filter(Boolean);
  if(!headers.length) return false;
  const hasPendiente = headers.includes('PENDIENTE');
  const hasCliente = headers.includes('CLIENTE');
  const hasNumero = headers.includes('NUMERO PENDIENTE');
  const hasComercial = headers.includes('COMERCIAL');
  const hasCategoria = headers.includes('CATEGORIA');
  const hasValor = headers.includes('VALOR FACTURAS');
  return hasPendiente && (hasCliente || hasNumero || hasComercial || hasCategoria || hasValor);
}

function findHeaderRowIndex(rows){
  for(let i = 0; i < rows.length; i++) {
    if(isLikelyHeaderRow(rows[i])) return i;
  }
  return -1;
}

function findSalesPendingHeaderRowIndex(rows){
  for(let i = 0; i < rows.length; i++) {
    if(isLikelySalesPendingHeaderRow(rows[i])) return i;
  }
  return -1;
}

function isMeaningfulDataRow(row){
  return Array.isArray(row) && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '');
}

function hasMeaningfulRecordContent(rec){
  if(!rec || typeof rec !== 'object') return false;

  const textKeys = [
    'CLIENTE','NOMBRE CLIENTE','EMPRESA','RAZON SOCIAL',
    'PRODUCTO','PROYECTO','SOLUCION','SERVICIO',
    'NUMERO DE PARTE','NUMERO PARTE','NRO PARTE','NO PARTE','PART NUMBER',
    'NUMERO DE COTIZACION','NUMERO COTIZACION','NRO COTIZACION','NO COTIZACION','COTIZACION',
    'MARCA','FABRICANTE','LINEA DE PRODUCTO','LINEA',
    'OBSERVACIONES','OBSERVACION','COMENTARIOS','COMENTARIO'
  ];

  if(textKeys.some(key => cleanDisplayText(rec[key], '') !== '')) return true;

  const numericKeys = ['MONTO VENTA CLIENTE','VALOR VENTA','MONTO VENTA','COSTO NEGOCIO','COSTO','UTILIDAD'];
  if(numericKeys.some(key => (parseMonto(rec[key]) || 0) !== 0)) return true;

  const dateKeys = ['FECHA DIA/MES/AÑO','FECHA','FECHA NEGOCIO','FECHA VENTA'];
  if(dateKeys.some(key => cleanDisplayText(rec[key], '') !== '')) return true;

  return false;
}

function hasMeaningfulSalesPendingRecordContent(rec){
  if(!rec || typeof rec !== 'object') return false;
  const textKeys = ['PENDIENTE','NUMERO PENDIENTE','CLIENTE','COMERCIAL','CATEGORIA'];
  if(textKeys.some(key => cleanDisplayText(rec[key], '') !== '')) return true;
  return (parseMonto(rec['VALOR FACTURAS']) || 0) !== 0;
}

function parseSalesSupportFileName(name){
  const base = String(name || '').replace(/\.(xlsx|xls)$/i,'').trim();
  if(!isSalesSupportFile(base)) return null;
  const rest = base
    .replace(/^sales[\s._\-–—]*support/i,'')
    .replace(/^[-–—_\s]+/,'')
    .trim();
  const parts = rest.split(/\s+-\s+/).map(cleanNameSegment).filter(Boolean);
  const soportaNames = parts.slice(1);
  return {
    supportName: cleanNameSegment(parts[0] || ''),
    soportaNames,
    soportaName: formatSalesTargets(soportaNames)
  };
}

function decorateRecordFromFile(rec, fileName, directorHint){
  const salesMeta = parseSalesSupportFileName(fileName);
  rec['CLIENTE'] = firstFilled(rec, ['CLIENTE']) || rec['CLIENTE'] || '';
  rec['NUMERO DE COTIZACION'] = firstFilled(rec, ['NUMERO DE COTIZACION']) || rec['NUMERO DE COTIZACION'] || '';
  rec['COSTO NEGOCIO'] = firstFilled(rec, ['COSTO NEGOCIO','COSTO']) || rec['COSTO NEGOCIO'] || '';
  rec['MONTO VENTA CLIENTE'] = firstFilled(rec, ['MONTO VENTA CLIENTE']) || rec['MONTO VENTA CLIENTE'] || '';
  rec['UTILIDAD'] = firstFilled(rec, ['UTILIDAD']) || rec['UTILIDAD'] || '';
  rec['MARGEN'] = firstFilled(rec, ['MARGEN']) || rec['MARGEN'] || '';
  rec['MONEDA 2'] = firstFilled(rec, ['MONEDA 2']) || rec['MONEDA 2'] || '';
  rec['TRM REFERENCIA'] = firstFilled(rec, ['TRM REFERENCIA']) || rec['TRM REFERENCIA'] || '';
  rec['FECHA DIA/MES/AÑO'] = firstFilled(rec, ['FECHA DIA/MES/AÑO']) || rec['FECHA DIA/MES/AÑO'] || '';
  rec['DIRECTOR'] = normalizeDirectorName(directorHint || rec['DIRECTOR'] || rec['DIRECTOR '] || '');
  rec['ESTADO'] = normalizeEstado(firstFilled(rec, ['ESTADO']));
  if(salesMeta){
    const supportName = canonicalizeSalesSupportName(firstFilled(rec, ['SALES SUPPORT']) || salesMeta.supportName);
    const soportaName = formatSalesTargets(firstFilled(rec, ['SOPORTA']) || salesMeta.soportaNames);
    rec['SALES SUPPORT'] = supportName;
    rec['COMERCIAL'] = supportName || cleanNameSegment(rec['COMERCIAL']);
    rec['SOPORTA'] = soportaName;
  } else {
    const fileNameExec = toTitleName(fileName.replace(/\.(xlsx|xls)$/i,'').trim());
    rec['COMERCIAL'] = fileNameExec || toTitleName(rec['COMERCIAL'] || '');
  }
  return rec;
}

function normalizeSalesPendingType(value){
  const raw = cleanDisplayText(value, '');
  const key = normalizeCategoryValue(raw);
  if(!key) return '';
  if(key.includes('pedido')) return 'Pedido';
  if(key.includes('glpi')) return 'GLPI';
  if(key.includes('remision') || key.includes('remicion') || key.includes('reminsion')) return 'Remision';
  return toTitleName(raw);
}

function normalizeSalesPendingCategory(value, pendingType){
  const raw = cleanDisplayText(value, '');
  const key = normalizeCategoryValue(raw);
  if(key.includes('compra')) return 'Compra';
  if(key.includes('garantia')) return 'Garantias';
  if(key.includes('factura') || key.includes('facturado')) return 'Factura';

  const typeKey = normalizeCategoryValue(pendingType);
  if(typeKey.includes('pedido')) return 'Compra';
  if(typeKey.includes('glpi')) return 'Garantias';
  if(typeKey.includes('remision') || typeKey.includes('remicion') || typeKey.includes('reminsion')) return 'Factura';

  return raw ? toTitleName(raw) : '';
}

function decorateSalesPendingRecordFromFile(rec, fileName, directorHint){
  const salesMeta = parseSalesSupportFileName(fileName);
  const pendingType = normalizeSalesPendingType(firstFilled(rec, ['PENDIENTE']));
  const commercial = cleanNameSegment(
    firstFilled(rec, ['COMERCIAL','SOPORTA'])
    || (salesMeta && salesMeta.soportaName)
    || ''
  );
  const supportName = canonicalizeSalesSupportName(
    firstFilled(rec, ['SALES SUPPORT'])
    || (salesMeta && salesMeta.supportName)
    || ''
  );

  rec['PENDIENTE'] = pendingType;
  rec['NUMERO PENDIENTE'] = firstFilled(rec, ['NUMERO PENDIENTE']) || rec['NUMERO PENDIENTE'] || '';
  rec['CLIENTE'] = firstFilled(rec, ['CLIENTE']) || rec['CLIENTE'] || '';
  rec['COMERCIAL'] = commercial;
  rec['SOPORTA'] = commercial;
  rec['CATEGORIA'] = normalizeSalesPendingCategory(firstFilled(rec, ['CATEGORIA']), pendingType);
  rec['VALOR FACTURAS'] = firstFilled(rec, ['VALOR FACTURAS']) || rec['VALOR FACTURAS'] || '';
  rec['SALES SUPPORT'] = supportName;
  rec['DIRECTOR'] = normalizeDirectorName(directorHint || rec['DIRECTOR'] || rec['DIRECTOR '] || '');
  return rec;
}

function sumUtilidad(data, moneda){
  return data.reduce((acc, row) => {
    const utilidad = getUtilidad(row);
    return acc + (utilidad.moneda === moneda ? utilidad.valor : 0);
  }, 0);
}

/* Get month from date string */
function getMonth(d){
  return parsefecha(d);
}

function getForecastMonths(data, year){
  const allMonths = [...new Set((data || []).map(r => getMonth(getRowDateValue(r))).filter(m => /^\d{4}-\d{2}$/.test(m)))].sort();
  const targetYear = String(year || FORECAST_YEAR || '').trim();
  const yearMonths = targetYear ? allMonths.filter(month => month.startsWith(targetYear + '-')) : allMonths;
  return yearMonths.length ? yearMonths : allMonths;
}

function getMonthIndex(monthKey){
  const parts = String(monthKey || '').split('-');
  const idx = Number(parts[1]) - 1;
  return idx >= 0 && idx < 12 ? idx : -1;
}

function getMonthShortLabel(monthKey){
  const idx = getMonthIndex(monthKey);
  return idx >= 0 ? MONTH_LABELS_SHORT[idx] : monthKey;
}

function getMonthLongLabel(monthKey){
  const idx = getMonthIndex(monthKey);
  return idx >= 0 ? MONTH_LABELS_LONG[idx] : monthKey;
}

function syncMonthSelectOptions(selectId, months){
  const sel = document.getElementById(selectId);
  if(!sel) return '';
  const current = sel.value;
  sel.innerHTML = optionHtml('', 'Todos', false) + buildOptionList(months, {
    getLabel: getMonthLongLabel
  });
  if(current && months.includes(current)) sel.value = current;
  else sel.value = '';
  return sel.value;
}

function refreshForecastMonthFilters(){
  const months = getForecastMonths(getVisibleData());
  syncMonthSelectOptions('sel-dir-mes', months);
  syncMonthSelectOptions('sel-ej-mes', months);
}

/* Today string */
function todayStr(){
  const n = new Date();
  return n.toISOString().substring(0,10);
}

/* ══════════════════════════════════════
   FILE LOADING
══════════════════════════════════════ */

// Extrae el nombre del director desde el path de la carpeta
// Ej: "FORECAST 2026/Grupo Juan David Novoa/Freddy.xlsx" → "Juan David Novoa"
function directorFromPath(path){
  const parts = path.split('/');
  // Buscar la parte que empiece con "Grupo" o "Gupo" (typo en sharepoint)
  const dirPart = parts.find(p=> /^(Grupo|Gupo)/i.test(p.trim()));
  if(dirPart){
    return normalizeDirectorName(dirPart.replace(/^(Grupo|Gupo)\s+/i,'').trim());
  }
  // Si no, usar el folder padre del archivo
  if(parts.length>=2) return normalizeDirectorName(parts[parts.length-2].trim());
  return '';
}

// Archivos cargados por director, incluyendo archivos sin filas válidas
let LOADED_FILES_BY_DIR = {};

// TRM change
document.getElementById('trm-input').addEventListener('input',function(){
  TRM_READY = true;
  TRM=getTRM();
  if(ALL_DATA.length || SALES_DATA.length || SALES_PENDING_DATA.length || PREVENTA_DATA.length) renderVisiblePage();
});

function finalizeLoad(){
  // Ocultar zona de carga, mostrar solo botón reload
  const uz = document.getElementById('upload-zone-g');
  if(uz) uz.style.display = 'none';
  const reloadBar = document.getElementById('reload-bar');
  if(reloadBar) reloadBar.style.display = 'flex';
  const reloadInfo = document.getElementById('reload-info');
  if(reloadInfo) {
    const nFiles = Object.values(LOADED_FILES_BY_DIR).reduce((s,a)=>s+a.length,0);
    const nSalesFiles = Object.values(LOADED_SALES_BY_SUPPORT).reduce((s,a)=>s+a.length,0);
    const nPreventaFiles = LOADED_PREVENTA_FILES.length;
    const nDirs = Object.keys(LOADED_FILES_BY_DIR).length;
    reloadInfo.textContent = CURRENT_USER && CURRENT_USER.role === 'sales_support'
      ? nSalesFiles + ' archivos sales · Última carga: ' + new Date().toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'})
      : nFiles + ' archivos cargados · ' + nDirs + ' equipos' + (nSalesFiles ? ' · '+nSalesFiles+' sales' : '') + (nPreventaFiles ? ' · '+nPreventaFiles+' preventa' : '') + ' · Última carga: ' + new Date().toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'});
  }
  // TRM se mantiene desde Banrep (no se sobrescribe con Excel)
  TRM=getTRM();
  // Asegurar visibilidad correcta sin importar el flujo que llamó
  const gc = document.getElementById('gerencia-content');
  if(gc) gc.style.display='block';
  const hs = document.getElementById('hoy-strip');
  if(hs) hs.style.display='block';

  // Status header
  const visibleData = getVisibleData();
  const visibleSales = getVisibleSalesData();
  const visibleSalesPending = getVisibleSalesPendingData();
  const visiblePreventa = getVisiblePreventaData();
  const dirs = [...new Set(visibleData.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))].sort();
  const execs = [...new Set(visibleData.map(r=>r['COMERCIAL']||'').filter(Boolean))].sort();
  const execsWithData = [...new Set(visibleData.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  const salesSupports = [...new Set([
    ...visibleSales.map(r=>getSalesSupportName(r)).filter(Boolean),
    ...visibleSalesPending.map(r=>getSalesPendingSupportName(r)).filter(Boolean)
  ])];
  const salesTargets = [...new Set(visibleSales.map(r=>getSalesSoportaName(r)).filter(Boolean))];
  document.getElementById('file-count-hd').textContent =
    CURRENT_USER && CURRENT_USER.role === 'sales_support'
      ? `${visibleSales.length} registros sales · ${visibleSalesPending.length} pendientes · ${salesSupports.length} support · ${salesTargets.length} apoyos`
      : `${visibleData.length} negocios · ${dirs.length} dir · ${execsWithData.length} ejecutivos${visibleSales.length ? ' · '+visibleSales.length+' sales' : ''}${visibleSalesPending.length ? ' · '+visibleSalesPending.length+' pendientes' : ''}${visiblePreventa.length ? ' · '+visiblePreventa.length+' preventa' : ''}`;
  const now = new Date();
  document.getElementById('last-update-hd').textContent=
    'Actualizado: '+now.toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'})+' · '+
    now.toLocaleDateString('es-CO',{day:'2-digit',month:'short'});
  const dot = document.getElementById('status-dot-hd');
  if(dot){dot.style.background='var(--corp-green)';dot.style.boxShadow='0 0 8px var(--corp-green)';}
  const reloadBtn = document.getElementById('btn-reload');
  if(reloadBtn){
    reloadBtn.style.background='rgba(13,191,130,.15)';
    reloadBtn.style.border='1px solid rgba(13,191,130,.35)';
    reloadBtn.style.color='var(--corp-green)';
    const lbl=document.getElementById('btn-reload-txt');
    if(lbl) lbl.textContent='🔄 Recargar';
  }
  // Populate selects
  const selDir=document.getElementById('sel-director');
  const dirsForSel=[...new Set([
    ...visibleData.map(r=>(r['DIRECTOR']||'').trim()),
    ...ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()),
    ...Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim())
  ].filter(Boolean))].sort();
  selDir.innerHTML=buildOptionList(dirsForSel);
  
  const execsFromFiles2=Object.values(LOADED_FILES_BY_DIR||{}).flat()
    .map(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim()).filter(Boolean);
  const allExecsForSel=[...new Set([...execs,...execsFromFiles2])].sort();
  const selEj=document.getElementById('sel-ejecutivo');
  selEj.innerHTML=buildOptionList(allExecsForSel);
  refreshForecastMonthFilters();

  // Aplicar pestañas y badge según rol una vez que los selects ya están listos
  applyRoleTabs();
  showUserBadge();
}

function getVisibleData() {
  if(!CURRENT_USER) return ALL_DATA;
  const { role, directorGroup, name } = CURRENT_USER;
  if(role === 'sales_support') return [];
  if(role === 'director') {
    const targetDirector = normalizePersonName(directorGroup);
    return ALL_DATA.filter(r => normalizePersonName(r['DIRECTOR']) === targetDirector);
  }
  if(role === 'ejecutivo') {
    const targetName = getExecTargetName();
    const targetNorm = normalizePersonName(targetName);
    return ALL_DATA.filter(r => {
      const execName = (r['COMERCIAL']||'').trim();
      const execNorm = normalizePersonName(execName);
      return execNorm === targetNorm || namesMatch(execName, targetName);
    });
  }
  return ALL_DATA; // gerencia ve todo
}

function hasGerenciaCrossfilters(){
  return Object.values(GERENCIA_CROSSFILTERS).some(Boolean);
}

function applyGerenciaCrossfilters(rows){
  const f = GERENCIA_CROSSFILTERS;
  return (rows || []).filter(row => {
    if(f.estado && cleanDisplayText(row['ESTADO'],'').toUpperCase() !== f.estado) return false;
    if(f.director && normalizePersonName(row['DIRECTOR']) !== normalizePersonName(f.director)) return false;
    if(f.ejecutivo && !namesMatch(row['COMERCIAL'] || '', f.ejecutivo)) return false;
    if(f.linea && getRowLineName(row) !== f.linea) return false;
    return true;
  });
}

function setGerenciaCrossfilter(type, value){
  if(!Object.prototype.hasOwnProperty.call(GERENCIA_CROSSFILTERS, type)) return;
  const next = cleanDisplayText(value, '');
  GERENCIA_CROSSFILTERS[type] = GERENCIA_CROSSFILTERS[type] === next ? '' : next;
  renderGerencia();
}

function clearGerenciaCrossfilter(type){
  if(type && Object.prototype.hasOwnProperty.call(GERENCIA_CROSSFILTERS, type)) {
    GERENCIA_CROSSFILTERS[type] = '';
  } else {
    GERENCIA_CROSSFILTERS = { estado:'', director:'', ejecutivo:'', linea:'' };
  }
  renderGerencia();
}

function renderGerenciaCrossfilterBar(totalRows, visibleRows){
  const el = document.getElementById('gerencia-crossfilters');
  if(!el) return;
  const labels = [
    ['estado','Estado'],
    ['director','Director'],
    ['ejecutivo','Ejecutivo'],
    ['linea','Linea']
  ];
  const chips = labels
    .filter(([key]) => GERENCIA_CROSSFILTERS[key])
    .map(([key,label]) => `<button type="button" class="crossfilter-chip" onclick="${escAttr(jsCall('clearGerenciaCrossfilter', key))}" title="Quitar filtro ${escAttr(label)}">
      <span>${escHtml(label)}</span>${escHtml(GERENCIA_CROSSFILTERS[key])}<strong>×</strong>
    </button>`)
    .join('');
  if(!chips) {
    el.innerHTML = '';
    return;
  }
  el.innerHTML = `
    <div class="crossfilter-summary">${fmtNum(visibleRows)} de ${fmtNum(totalRows)} registros</div>
    <div class="crossfilter-chips">${chips}</div>
    <button type="button" class="crossfilter-clear" onclick="clearGerenciaCrossfilter()">Limpiar filtros</button>
  `;
}

function getVisualCrossfilters(scope){
  return VISUAL_CROSSFILTERS[scope] || {};
}

function setVisualCrossfilter(scope, type, value){
  const filters = VISUAL_CROSSFILTERS[scope];
  if(!filters || !Object.prototype.hasOwnProperty.call(filters, type)) return;
  const next = cleanDisplayText(value, '');
  filters[type] = filters[type] === next ? '' : next;
  renderPage(scope);
}

function clearVisualCrossfilter(scope, type){
  const filters = VISUAL_CROSSFILTERS[scope];
  if(!filters) return;
  if(type && Object.prototype.hasOwnProperty.call(filters, type)) {
    filters[type] = '';
  } else {
    Object.keys(filters).forEach(key => { filters[key] = ''; });
  }
  renderPage(scope);
}

function buildVisualCrossfilterBar(scope, totalRows, visibleRows, labels){
  const filters = getVisualCrossfilters(scope);
  const chips = (labels || [])
    .filter(([key]) => filters[key])
    .map(([key,label]) => `<button type="button" class="crossfilter-chip" onclick="${escAttr(jsCall('clearVisualCrossfilter', scope, key))}" title="Quitar filtro ${escAttr(label)}">
      <span>${escHtml(label)}</span>${escHtml(filters[key])}<strong>×</strong>
    </button>`)
    .join('');
  if(!chips) return '';
  return `<div class="crossfilter-bar">
    <div class="crossfilter-summary">${fmtNum(visibleRows)} de ${fmtNum(totalRows)} registros</div>
    <div class="crossfilter-chips">${chips}</div>
    <button type="button" class="crossfilter-clear" onclick="${escAttr(jsCall('clearVisualCrossfilter', scope))}">Limpiar filtros</button>
  </div>`;
}

function toggleDirectorEstadoFilter(value){
  const sel = document.getElementById('sel-dir-estado');
  if(sel) sel.value = sel.value === value ? '' : value;
  renderDirector();
}

function toggleDirectorMonthFilter(value){
  const sel = document.getElementById('sel-dir-mes');
  if(sel) sel.value = sel.value === value ? '' : value;
  renderDirector();
}

function toggleEjecutivoEstadoFilter(value){
  const sel = document.getElementById('sel-ej-estado');
  if(sel) sel.value = sel.value === value ? '' : value;
  renderEjecutivo();
}

function toggleSalesEstadoFilter(value){
  const sel = document.getElementById('sel-sales-estado');
  if(sel) sel.value = sel.value === value ? '' : value;
  renderSales();
}

function toggleSalesPendingFilter(kind, value){
  const sel = kind === 'categoria'
    ? document.getElementById('sel-sales-categoria')
    : document.getElementById('sel-sales-pendiente');
  const next = sel && sel.value === value ? '' : value;
  setSalesPendingFilter(kind, next);
}

function togglePreventaEstadoFilter(value){
  const sel = document.getElementById('sel-preventa-estado');
  if(sel) sel.value = sel.value === value ? '' : value;
  renderPreventa();
}

function isOscarMarcasGlobalScope(){
  if(!CURRENT_USER || CURRENT_USER.role !== 'director') return false;
  return normalizePersonName(CURRENT_USER.directorGroup) === normalizePersonName('Oscar Beltran');
}

function getVisibleMarcasData(){
  if(isOscarMarcasGlobalScope()) return ALL_DATA;
  return getVisibleData();
}

function getSalesSupportTargetName() {
  const email = (CURRENT_USER && CURRENT_USER.email || '').toLowerCase().trim();
  const map = window.SALES_SUPPORT_BY_EMAIL || {};
  return canonicalizeSalesSupportName((map[email] || CURRENT_USER && CURRENT_USER.name || '').trim());
}

function getVisibleSalesData() {
  if(!CURRENT_USER) return SALES_DATA;
  const { role, directorGroup } = CURRENT_USER;
  if(role === 'sales_support') {
    const targetName = getSalesSupportTargetName();
    return SALES_DATA.filter(r => namesMatch(getSalesSupportName(r), targetName));
  }
  if(role === 'director') {
    const targetDirector = normalizePersonName(directorGroup);
    return SALES_DATA.filter(r => normalizePersonName(r['DIRECTOR']) === targetDirector);
  }
  if(role === 'ejecutivo') return [];
  return SALES_DATA;
}

function getVisibleSalesPendingData() {
  if(!CURRENT_USER) return SALES_PENDING_DATA;
  const { role, directorGroup } = CURRENT_USER;
  if(role === 'sales_support') {
    const targetName = getSalesSupportTargetName();
    return SALES_PENDING_DATA.filter(r => namesMatch(getSalesPendingSupportName(r), targetName));
  }
  if(role === 'director') {
    const targetDirector = normalizePersonName(directorGroup);
    return SALES_PENDING_DATA.filter(r => normalizePersonName(r['DIRECTOR']) === targetDirector);
  }
  if(role === 'ejecutivo') return [];
  return SALES_PENDING_DATA;
}

function getPreventaName(row){
  return firstFilled(row, ['PREVENTA','COMERCIAL']) || '';
}

function getVisiblePreventaData() {
  if(!CURRENT_USER) return PREVENTA_DATA;
  if(CURRENT_USER.role === 'sales_support' || CURRENT_USER.role === 'ejecutivo' || CURRENT_USER.role === 'director') return [];
  return PREVENTA_DATA;
}

function renderAll(){
  refreshForecastMonthFilters();
  renderGerencia();
  renderDirector();
  renderEjecutivo();
  renderSales();
  renderPreventa();
  renderDivisas();
  renderMarcas();
  renderResumen();
  if(MARCA_LINEA_DETAIL_STATE && document.getElementById('page-marca-linea-detail')) {
    renderMarcaLineaDetail();
  }
  if(NEGOCIO_DETAIL_STATE && document.getElementById('page-negocio')) {
    const row = getRecordById(NEGOCIO_DETAIL_STATE.rowId);
    if(row) renderNegocioDetail(row);
  }
}

function renderPage(pageId){
  const page = pageId || getActivePageId();
  if(page === 'gerencia') {
    renderGerencia();
    return;
  }
  if(page === 'director') {
    renderDirector();
    return;
  }
  if(page === 'ejecutivo') {
    renderEjecutivo();
    return;
  }
  if(page === 'sales') {
    renderSales();
    return;
  }
  if(page === 'preventa') {
    renderPreventa();
    return;
  }
  if(page === 'divisas') {
    renderDivisas();
    return;
  }
  if(page === 'marcas') {
    renderMarcas();
    return;
  }
  if(page === 'resumen') {
    renderResumen();
    return;
  }
  if(page === 'finanzas') {
    renderFinanzas();
    return;
  }
  if(page === 'marca-linea-detail' && MARCA_LINEA_DETAIL_STATE) {
    renderMarcaLineaDetail();
    return;
  }
  if(page === 'negocio' && NEGOCIO_DETAIL_STATE) {
    const row = getRecordById(NEGOCIO_DETAIL_STATE.rowId);
    if(row) renderNegocioDetail(row);
  }
}

function renderVisiblePage(){
  refreshForecastMonthFilters();
  renderPage(getActivePageId());
}

/* ══════════════════════════════════════
   NAV
══════════════════════════════════════ */
function showPage(id,btn){
  const currentPage = getActivePageId();
  if(currentPage === 'negocio' && id !== 'negocio') NEGOCIO_DETAIL_STATE = null;
  if(currentPage === 'marca-linea-detail' && id !== 'marca-linea-detail' && id !== 'negocio') MARCA_LINEA_DETAIL_STATE = null;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  if(btn) btn.classList.add('active');
  const hasLoadedData = ALL_DATA.length || SALES_DATA.length || SALES_PENDING_DATA.length || PREVENTA_DATA.length;
  if(hasLoadedData || id === 'finanzas' || id === 'negocio' || id === 'marca-linea-detail') {
    renderPage(id);
  }
}

function getActivePageId(){
  const active = document.querySelector('.page.active');
  return active ? active.id.replace('page-','') : 'gerencia';
}

function getScrollTrackersForPage(page){
  const pageEl = document.getElementById('page-' + page);
  if(!pageEl) return [];
  return [...pageEl.querySelectorAll('[data-scroll-key]')].map(el => ({
    key: el.getAttribute('data-scroll-key'),
    el
  }));
}

function captureScrollSnapshot(page){
  return {
    windowY: window.scrollY || window.pageYOffset || 0,
    nodes: getScrollTrackersForPage(page).reduce((acc, item) => {
      acc[item.key] = {
        top: item.el.scrollTop || 0,
        left: item.el.scrollLeft || 0
      };
      return acc;
    }, {})
  };
}

function restoreScrollSnapshot(page, snapshot){
  if(!snapshot) return;
  const apply = () => {
    window.scrollTo({ top: snapshot.windowY || 0, left: 0, behavior: 'auto' });
    getScrollTrackersForPage(page).forEach(({ key, el }) => {
      const pos = snapshot.nodes && snapshot.nodes[key];
      if(!pos) return;
      el.scrollTop = pos.top || 0;
      el.scrollLeft = pos.left || 0;
    });
  };
  requestAnimationFrame(() => {
    apply();
    requestAnimationFrame(apply);
  });
}

function getNavButtonForPage(page){
  return document.getElementById('tab-' + page);
}

function showMoreEstadoRows(estado){
  const scrollState = captureScrollSnapshot('gerencia');
  GERENCIA_ESTADO_LIMITS[estado] = (GERENCIA_ESTADO_LIMITS[estado] || 30) + GERENCIA_ESTADO_STEP;
  renderGerenciaEstadoTables(applyGerenciaCrossfilters(getVisibleData()));
  restoreScrollSnapshot('gerencia', scrollState);
}

function setDivisaEstadoFilter(value){
  const sel = document.getElementById('sel-divisa-estado');
  if(sel) sel.value = value;
  DIVISAS_DETAIL_LIMITS.COP = 10;
  DIVISAS_DETAIL_LIMITS.USD = 10;
  renderDivisas();
}

function showMoreDivisaRows(moneda){
  const key = (moneda || '').toUpperCase();
  if(!DIVISAS_DETAIL_LIMITS[key]) return;
  const scrollState = captureScrollSnapshot('divisas');
  DIVISAS_DETAIL_LIMITS[key] = (DIVISAS_DETAIL_LIMITS[key] || 10) + DIVISAS_DETAIL_STEP;
  renderDivisas();
  restoreScrollSnapshot('divisas', scrollState);
}

function showMoreMarcasBars(){
  const scrollState = captureScrollSnapshot('marcas');
  MARCAS_BAR_LIMIT = (MARCAS_BAR_LIMIT || MARCAS_BAR_INITIAL) + MARCAS_BAR_STEP;
  renderMarcas();
  restoreScrollSnapshot('marcas', scrollState);
}

function normalizeCategoryValue(value){
  return cleanDisplayText(value, '')
    .toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function getMarcaLineaDetailTypeLabel(type){
  return ({ marca:'Marca', linea:'Linea', producto:'Producto' })[type] || 'Categoria';
}

function getMarcaLineaDetailRowValue(row, type){
  if(type === 'marca') return getRowBrandName(row);
  if(type === 'linea') return getRowLineName(row);
  if(type === 'producto') return getRowProductName(row);
  return '';
}

function getMarcaLineaDetailRows(state){
  if(!state || !state.value) return [];
  const target = normalizeCategoryValue(state.value);
  return getVisibleMarcasData().filter(row => normalizeCategoryValue(getMarcaLineaDetailRowValue(row, state.type)) === target);
}

function openMarcaLineaDetail(type, value, sourcePage){
  const cleanValue = cleanDisplayText(value, '').trim();
  if(!cleanValue) return;
  const backPage = sourcePage || getActivePageId();
  MARCA_LINEA_DETAIL_STATE = {
    type,
    value: cleanValue,
    estadoFilter: '',
    backPage,
    backScroll: captureScrollSnapshot(backPage)
  };
  showPage('marca-linea-detail');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function setMarcaLineaDetailEstado(value){
  if(!MARCA_LINEA_DETAIL_STATE) return;
  MARCA_LINEA_DETAIL_STATE.estadoFilter = value || '';
  renderMarcaLineaDetail();
}

function closeMarcaLineaDetail(){
  const backPage = (MARCA_LINEA_DETAIL_STATE && MARCA_LINEA_DETAIL_STATE.backPage) || 'marcas';
  const backScroll = MARCA_LINEA_DETAIL_STATE && MARCA_LINEA_DETAIL_STATE.backScroll;
  MARCA_LINEA_DETAIL_STATE = null;
  showPage(backPage, getNavButtonForPage(backPage));
  restoreScrollSnapshot(backPage, backScroll);
}

function openNegocioDetailById(id, sourcePage){
  const row = getRecordById(id);
  if(!row) return;
  const backPage = sourcePage || getActivePageId();
  NEGOCIO_DETAIL_STATE = {
    rowId: id,
    backPage,
    backScroll: captureScrollSnapshot(backPage)
  };
  showPage('negocio');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function closeNegocioDetail(){
  const backPage = (NEGOCIO_DETAIL_STATE && NEGOCIO_DETAIL_STATE.backPage) || 'gerencia';
  const backScroll = NEGOCIO_DETAIL_STATE && NEGOCIO_DETAIL_STATE.backScroll;
  NEGOCIO_DETAIL_STATE = null;
  showPage(backPage, getNavButtonForPage(backPage));
  restoreScrollSnapshot(backPage, backScroll);
}

function formatFieldValue(key, value, row){
  const mapped = normalizeHeaderKey(mapHeaderName(key));
  if(value === null || value === undefined || String(value).trim() === '') return 'Sin dato';
  if(mapped.startsWith('FECHA')) return escHtml(cleanDisplayText(formatDateValue(value), 'Sin fecha'));
  if(mapped === 'MONTO VENTA CLIENTE' || mapped === 'COSTO NEGOCIO' || mapped === 'COSTO' || mapped === 'UTILIDAD' || mapped === 'VALOR FACTURAS'){
    const mon = cleanDisplayText(row['MONEDA 2'], 'COP').toUpperCase();
    const monto = parseMonto(value);
    return mapped !== 'VALOR FACTURAS' && mon === 'USD' ? fmtUSD(monto) : fmtCOP(monto);
  }
  if(mapped === 'TRM REFERENCIA') return fmtTRM(getTRM());
  if(mapped === 'MARGEN') return escHtml(formatMarginDisplay(value));
  if(value instanceof Date) return escHtml(cleanDisplayText(formatDateValue(value), 'Sin fecha'));
  if(typeof value === 'number') return escHtml(value.toLocaleString('es-CO'));
  return escHtml(cleanDisplayText(value, 'Sin dato'));
}

function renderNegocioDetail(row){
  const host = document.getElementById('negocio-detail');
  if(!host || !row) return;

  const isSalesRow = row.__DATASET === 'sales';
  const estado = cleanDisplayText(row['ESTADO'], 'Sin estado').toUpperCase();
  const estadoClass = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'].includes(estado) ? estado : 'PENDIENTE';
  const cliente = cleanDisplayText(getRowClientName(row), 'Sin cliente');
  const producto = cleanDisplayText(getRowProductName(row), isSalesRow ? 'Registro Sales Support' : 'Sin proyecto');
  const marca = cleanDisplayText(getRowBrandName(row), 'Sin marca');
  const numeroParte = isSalesRow
    ? cleanDisplayText(getSalesQuoteNumber(row), 'Sin numero de cotizacion')
    : cleanDisplayText(getRowPartNumber(row), 'Sin numero de parte');
  const linea = cleanDisplayText(getRowLineName(row), 'Sin linea');
  const soporta = cleanDisplayText(getSalesSoportaName(row), 'Sin apoyo definido');
  const director = cleanDisplayText(firstFilled(row, ['DIRECTOR']), 'Sin director');
  const ejecutivo = cleanDisplayText(firstFilled(row, ['SALES SUPPORT','COMERCIAL']), isSalesRow ? 'Sin sales support' : 'Sin ejecutivo');
  const moneda = cleanDisplayText(firstFilled(row, ['MONEDA 2']), 'COP').toUpperCase();
  const valor = parseMonto(row['MONTO VENTA CLIENTE']) || 0;
  const copTotal = toCOP(row);
  const margenRaw = firstFilled(row, ['MARGEN']);
  const margen = formatMarginDisplay(margenRaw);
  const fecha = cleanDisplayText(formatDateValue(getRowDateValue(row)), 'Sin fecha');
  const trmRef = getTRM();
  const observacion = cleanDisplayText(getRowObservation(row), 'Sin observacion registrada en el Excel.');
  const modificacion = cleanDisplayText(firstFilled(row, ['MODIFICACION O UPGRADE']), 'Sin dato');
  const sourceFile = cleanDisplayText(firstFilled(row, ['__SOURCE_FILE']), 'Sin archivo');
  const sourceSheet = cleanDisplayText(firstFilled(row, ['__SOURCE_SHEET']), 'Sin hoja');
  const responsibleLabel = isSalesRow ? 'Sales Support' : 'Ejecutivo';
  const numberLabel = isSalesRow ? 'Cotizacion' : 'Parte';

  const orderedKeys = [
    'CLIENTE','NOMBRE CLIENTE','EMPRESA','PRODUCTO','PROYECTO','NUMERO DE PARTE','NUMERO DE COTIZACION',
    'MARCA','LINEA DE PRODUCTO','ESTADO','DIRECTOR','COMERCIAL','SALES SUPPORT','SOPORTA',
    'MONEDA 2','COSTO NEGOCIO','COSTO','MONTO VENTA CLIENTE','UTILIDAD','TRM REFERENCIA','FECHA DIA/MES/AÑO',
    'MARGEN','MODIFICACION O UPGRADE','OBSERVACIONES'
  ].map(normalizeHeaderKey);

  const allEntries = Object.entries(row)
    .filter(([key, value]) => {
      if(String(key).startsWith('__')) return false;
      if(value === null || value === undefined || String(value).trim() === '') return false;
      if(isSalesRow && normalizeHeaderKey(mapHeaderName(key)) === 'COMERCIAL' && row['SALES SUPPORT']) return false;
      return true;
    })
    .sort((a,b) => {
      const keyA = normalizeHeaderKey(mapHeaderName(a[0]));
      const keyB = normalizeHeaderKey(mapHeaderName(b[0]));
      const idxA = orderedKeys.indexOf(keyA);
      const idxB = orderedKeys.indexOf(keyB);
      if(idxA === -1 && idxB === -1) return formatFieldLabel(a[0]).localeCompare(formatFieldLabel(b[0]));
      if(idxA === -1) return 1;
      if(idxB === -1) return -1;
      return idxA - idxB;
    });

  host.innerHTML = `
    <div class="detail-shell">
      <div class="detail-topbar">
        <button type="button" class="btn-clear detail-back-btn" onclick="closeNegocioDetail()">Volver</button>
        <span class="section-tag">Vista individual del negocio</span>
      </div>

      <div class="detail-hero">
        <div>
          <div class="detail-overline">${isSalesRow ? 'Registro sales support' : 'Proyecto o negocio'}</div>
          <h2>${escHtml(producto)}</h2>
          <p>${isSalesRow ? `${escHtml(cliente)} / Soporta: ${escHtml(soporta)}` : `${escHtml(cliente)} / ${escHtml(marca)} / ${escHtml(linea)}`}</p>
          <div class="detail-chip-row">
            <span class="badge badge-${estadoClass}">${escHtml(estado)}</span>
            <span class="detail-chip">Director: ${escHtml(director)}</span>
            <span class="detail-chip">${escHtml(responsibleLabel)}: ${escHtml(ejecutivo)}</span>
            <span class="detail-chip">Fecha: ${escHtml(fecha)}</span>
            <span class="detail-chip">${escHtml(numberLabel)}: ${escHtml(numeroParte)}</span>
          </div>
        </div>
        <div class="detail-hero-amounts">
          <div>
            <div class="detail-amount-label">Valor negocio</div>
            <div class="detail-amount-main">${moneda === 'USD' ? fmtUSD(valor) : fmtCOP(valor)}</div>
          </div>
          <div>
            <div class="detail-amount-label">Total COP</div>
            <div class="detail-amount-alt">${fmtCOP(copTotal)}</div>
          </div>
        </div>
      </div>

      <div class="detail-kpi-grid">
        <div class="detail-stat">
          <span>${escHtml(numberLabel)}</span>
          <strong>${escHtml(numeroParte)}</strong>
          <small>${isSalesRow ? 'Identificador de la cotizacion' : 'Referencia del item o solucion'}</small>
        </div>
        <div class="detail-stat">
          <span>Moneda</span>
          <strong>${escHtml(moneda)}</strong>
          <small>TRM dia: ${fmtTRM(trmRef)}</small>
        </div>
        <div class="detail-stat">
          <span>Margen</span>
          <strong>${escHtml(margen)}</strong>
          <small>Upgrade: ${escHtml(modificacion)}</small>
        </div>
        <div class="detail-stat">
          <span>Cliente</span>
          <strong>${escHtml(cliente)}</strong>
          <small>Linea: ${escHtml(linea)}</small>
        </div>
        <div class="detail-stat">
          <span>Archivo fuente</span>
          <strong>${escHtml(sourceFile)}</strong>
          <small>Hoja: ${escHtml(sourceSheet)}</small>
        </div>
      </div>

      <div class="chart-card g1">
        <div class="chart-hd">Observacion comercial</div>
        <div class="detail-observation">${escHtml(observacion)}</div>
      </div>

      <div class="chart-card g1">
        <div class="chart-hd">Ficha completa del negocio</div>
        <div class="tbl-wrap">
          <table class="responsive-table responsive-table-detail">
            <thead><tr><th>Campo</th><th>Valor</th></tr></thead>
            <tbody>${allEntries.map(([key, value]) => `
              <tr>
                <td class="detail-field-name" data-label="Campo">${escHtml(formatFieldLabel(key))}</td>
                <td data-label="Valor">${formatFieldValue(key, value, row)}</td>
              </tr>
            `).join('')}</tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function scrollToElementAfterRender(id, offset){
  const topOffset = Number(offset) || 112;
  const apply = () => {
    const el = document.getElementById(id);
    if(!el) return;
    const top = el.getBoundingClientRect().top + (window.scrollY || window.pageYOffset || 0) - topOffset;
    window.scrollTo({ top: Math.max(top, 0), left: 0, behavior: 'smooth' });
  };
  requestAnimationFrame(() => requestAnimationFrame(apply));
}

function renderMarcaLineaDetail(){
  const host = document.getElementById('marca-linea-detail');
  const state = MARCA_LINEA_DETAIL_STATE;
  if(!host || !state) return;

  const typeLabel = getMarcaLineaDetailTypeLabel(state.type);
  const rows = getMarcaLineaDetailRows(state).sort((a,b)=>toCOP(b)-toCOP(a));
  const estadoFilter = state.estadoFilter || '';
  const filteredRows = estadoFilter ? rows.filter(r => cleanDisplayText(r['ESTADO'], '').toUpperCase() === estadoFilter) : rows;
  const totalCOP = rows.reduce((sum, row) => sum + toCOP(row), 0);
  const totalUSD = rows
    .filter(r => cleanDisplayText(r['MONEDA 2'], 'COP').trim().toUpperCase() === 'USD')
    .reduce((sum, row) => sum + (parseMonto(row['MONTO VENTA CLIENTE']) || 0), 0);
  const totalNegocios = rows.length;
  const estados = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'].map((estado) => {
    const estadoRows = rows.filter(r => cleanDisplayText(r['ESTADO'], '').toUpperCase() === estado);
    return {
      estado,
      count: estadoRows.length,
      total: estadoRows.reduce((sum, row) => sum + toCOP(row), 0)
    };
  });
  const topEstado = estados.slice().sort((a,b)=>b.total-a.total)[0] || { estado:'SIN DATOS', total:0, count:0 };

  host.innerHTML = `
    <div class="detail-shell">
      <div class="detail-topbar">
        <button type="button" class="btn-clear detail-back-btn" onclick="closeMarcaLineaDetail()">Volver</button>
        <span class="section-tag">Detalle de ${escHtml(typeLabel)}</span>
      </div>

      <div class="detail-hero">
        <div>
          <div class="detail-overline">${escHtml(typeLabel)} seleccionada</div>
          <h2>${escHtml(state.value)}</h2>
          <p>Vista consolidada de los negocios relacionados con esta ${escHtml(typeLabel.toLowerCase())}. Puedes revisar el comportamiento por estado y abrir cualquier negocio para ver su ficha completa.</p>
          <div class="detail-chip-row">
            <span class="detail-chip">Negocios: ${fmtNum(totalNegocios)}</span>
            <span class="detail-chip">Filtro actual: ${escHtml(estadoFilter || 'Todos')}</span>
            <span class="detail-chip">Estado lider: ${escHtml(topEstado.estado)}</span>
          </div>
        </div>
        <div class="detail-hero-amounts">
          <div>
            <div class="detail-amount-label">Total COP</div>
            <div class="detail-amount-main">${fmtCOP(totalCOP)}</div>
          </div>
          <div>
            <div class="detail-amount-label">Total USD</div>
            <div class="detail-amount-alt">${fmtUSD(totalUSD)}</div>
          </div>
        </div>
      </div>

      <div class="category-state-grid">
        <button type="button" class="category-state-card${!estadoFilter ? ' active' : ''}" onclick="setMarcaLineaDetailEstado('')">
          <span>Todos</span>
          <strong>${fmtNum(totalNegocios)}</strong>
          <small>${fmtCOP(totalCOP)}</small>
        </button>
        ${estados.map(item => `
          <button type="button" class="category-state-card state-${item.estado.toLowerCase()}${estadoFilter === item.estado ? ' active' : ''}" onclick="setMarcaLineaDetailEstado('${item.estado}')">
            <span>${escHtml(item.estado)}</span>
            <strong>${fmtNum(item.count)}</strong>
            <small>${fmtCOP(item.total)}</small>
          </button>
        `).join('')}
      </div>

      <div class="chart-card g1">
        <div class="director-table-toolbar">
          <div>
            <div class="chart-hd">Partes y clientes asociados</div>
            <div style="font-size:11px;color:var(--text2)">Agrupado por producto o numero de parte y cliente. La cantidad muestra cuantos registros tiene ese producto o cliente dentro de la seleccion y el valor se consolida en COP.</div>
          </div>
          <div class="director-table-filter">
            <div class="filter-label">Estado</div>
            <select id="sel-marca-linea-estado" onchange="setMarcaLineaDetailEstado(this.value)">
              <option value="">Todos</option>
              <option value="GANADA"${estadoFilter==='GANADA'?' selected':''}>GANADA</option>
              <option value="PENDIENTE"${estadoFilter==='PENDIENTE'?' selected':''}>PENDIENTE</option>
              <option value="PERDIDA"${estadoFilter==='PERDIDA'?' selected':''}>PERDIDA</option>
              <option value="APLAZADO"${estadoFilter==='APLAZADO'?' selected':''}>APLAZADO</option>
            </select>
          </div>
        </div>
        <div class="tbl-wrap">
          ${buildMarcaLineaDetailTable(filteredRows)}
        </div>
      </div>
    </div>
  `;
}

function buildMarcaLineaDetailTable(data){
  const grouped = new Map();

  (data || []).forEach(row => {
    const partNumber = cleanDisplayText(getRowPartNumber(row), '');
    const productName = cleanDisplayText(getRowProductName(row), '');
    const productKey = normalizeCategoryValue(partNumber || productName || 'sin numero de parte');
    const productLabel = partNumber || productName || 'Sin numero de parte';
    const clientLabel = cleanDisplayText(getRowClientName(row), 'Sin cliente');
    const clientKey = normalizeCategoryValue(clientLabel);
    const estado = cleanDisplayText(row['ESTADO'], 'Sin estado').toUpperCase();
    const key = [productKey, clientKey].join('||');
    const current = grouped.get(key) || {
      producto: productLabel,
      productoHint: productName,
      cliente: clientLabel,
      estados: {},
      cantidad: 0,
      totalCOP: 0,
      rows: []
    };

    current.cantidad += getRowQuantityValue(row, 1);
    current.totalCOP += toCOP(row);
    current.rows.push(row);
    current.estados[estado] = (current.estados[estado] || 0) + 1;
    if(!current.productoHint && productName) current.productoHint = productName;
    grouped.set(key, current);
  });

  const items = Array.from(grouped.values()).sort((a, b) =>
    (b.totalCOP - a.totalCOP) ||
    (b.cantidad - a.cantidad) ||
    a.cliente.localeCompare(b.cliente, 'es', { sensitivity: 'base' }) ||
    a.producto.localeCompare(b.producto, 'es', { sensitivity: 'base' })
  );

  return `<table class="responsive-table">
    <thead><tr><th>Producto</th><th>Cliente</th><th>Cantidad</th><th>Estado</th><th>Valor Total COP</th></tr></thead>
    <tbody>${items.length ? items.map(item => {
      const estadoEntries = Object.entries(item.estados || {}).sort((a,b)=>b[1]-a[1]);
      const singleRow = item.rows.length === 1 ? item.rows[0] : null;
      const rowAttrs = singleRow && singleRow.__RID
        ? ` class="table-row-action" onclick="${escAttr(jsCall('openNegocioDetailById', singleRow.__RID, 'marca-linea-detail'))}" title="Abrir detalle del negocio"`
        : '';
      const productTitle = cleanDisplayText(item.productoHint, '') || item.producto;
      return `<tr${rowAttrs}>
        <td style="color:var(--text);font-weight:600;max-width:180px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${escAttr(productTitle)}" data-label="Producto">${escHtml(item.producto)}</td>
        <td style="color:var(--text)" data-label="Cliente">${escHtml(item.cliente)}</td>
        <td class="td-mono" data-label="Cantidad">${fmtNum(item.cantidad)}</td>
        <td data-label="Estado">${estadoEntries.map(([estado,count]) => `<span class="badge badge-${getEstadoBadgeClass(estado)}" title="${escAttr(count + ' registros')}">${escHtml(estado)}${estadoEntries.length > 1 ? ' · ' + fmtNum(count) : ''}</span>`).join(' ')}</td>
        <td class="td-mono td-cop" data-label="Valor Total COP">${fmtCOP(item.totalCOP)}</td>
      </tr>`;
    }).join('') : `<tr><td colspan="5" style="text-align:center;color:var(--text2)">Sin registros para este filtro.</td></tr>`}</tbody>
  </table>`;
}

/* ══════════════════════════════════════
   CHART HELPERS
══════════════════════════════════════ */
function buildLineValueData(rows){
  const totals = new Map();
  (rows || []).forEach(row => {
    const linea = getRowLineName(row);
    if(!linea) return;
    totals.set(linea, (totals.get(linea) || 0) + toCOP(row));
  });
  return [...totals.entries()]
    .map(([name, val]) => ({ name, val }))
    .sort((a,b)=>b.val-a.val);
}

function buildBrandValueData(rows){
  const totals = new Map();
  (rows || []).forEach(row => {
    const marca = getRowBrandName(row);
    if(!marca) return;
    totals.set(marca, (totals.get(marca) || 0) + toCOP(row));
  });
  return [...totals.entries()]
    .map(([name, val]) => ({ name, val }))
    .sort((a,b)=>b.val-a.val);
}

function renderBars(containerId, items, color, fmtFn, opts){
  const el=document.getElementById(containerId);
  if(!el) return;
  const options = opts || {};
  const max=Math.max(...items.map(i=>i.val),1);
  const nameClass = options.nameClass || 'w100';
  el.innerHTML=items.map((it,idx)=>{
    const pct=Math.round((it.val/max)*100);
    const c=Array.isArray(color)?color[idx%color.length]:color;
    const onClick = options.getOnClick ? options.getOnClick(it, idx) : '';
    const isSelected = options.getIsSelected ? options.getIsSelected(it, idx) : false;
    const rowAttrs = onClick
      ? ` class="bar-row bar-row-action ${isSelected ? 'chart-selected' : ''}" role="button" tabindex="0" onclick="${escAttr(onClick)}" onkeydown="${escAttr(`if(event.key==='Enter'||event.key===' '){event.preventDefault();${onClick}}`)}" title="${escAttr(options.clickTitle || 'Filtrar')}" data-tooltip="${escAttr(options.tooltipPrefix ? options.tooltipPrefix + it.name : 'Filtrar por ' + it.name)}"`
      : ` class="bar-row ${isSelected ? 'chart-selected' : ''}"`;
    return `<div${rowAttrs}>
      <div class="bar-name ${escAttr(nameClass)}" title="${escAttr(it.name)}">${escHtml(it.name)}</div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${pct}%;background:${c}20;border:1px solid ${c}40">
          <span class="bar-val" style="color:${c}">${fmtFn?fmtFn(it.val):abr(it.val)}</span>
        </div>
      </div>
    </div>`;
  }).join('');
  attachChartTooltips(el);
}

function renderDonut(svgId, legId, items, opts){
  const svg=document.getElementById(svgId);
  const leg=document.getElementById(legId);
  if(!svg||!leg) return;
  const options = opts || {};
  const total=items.reduce((s,i)=>s+i.val,0)||1;
  let angle=-90;
  const r=38, cx=50, cy=50;
  let paths='';
  items.forEach((it,i)=>{
    const slice=(it.val/total)*360;
    const a1=angle*Math.PI/180;
    const a2=(angle+slice)*Math.PI/180;
    const x1=cx+r*Math.cos(a1), y1=cy+r*Math.sin(a1);
    const x2=cx+r*Math.cos(a2), y2=cy+r*Math.sin(a2);
    const large=slice>180?1:0;
    const c=COLORS[i%COLORS.length];
    const onClick = options.getOnClick ? options.getOnClick(it, i) : '';
    const isSelected = options.getIsSelected ? options.getIsSelected(it, i) : false;
    const activeMode = options.hasActive && !isSelected;
    const actionAttrs = onClick
      ? ` role="button" tabindex="0" onclick="${escAttr(onClick)}" onkeydown="${escAttr(`if(event.key==='Enter'||event.key===' '){event.preventDefault();${onClick}}`)}" data-tooltip="${escAttr(options.tooltipPrefix ? options.tooltipPrefix + it.name : 'Filtrar por ' + it.name)}"`
      : '';
    paths+=`<path class="donut-slice ${isSelected ? 'chart-selected' : ''}" d="M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} Z" fill="${c}" opacity="${activeMode ? '.35' : '.85'}" stroke="${isSelected ? c : 'transparent'}" stroke-width="${isSelected ? '2.2' : '0'}"${actionAttrs}/>`;
    angle+=slice;
  });
  svg.innerHTML=paths+`<circle cx="50" cy="50" r="22" fill="#0B0F1E"/>`;
  
  leg.innerHTML=items.slice(0,6).map((it,i)=>{
    const pct=((it.val/total)*100).toFixed(1);
    const onClick = options.getOnClick ? options.getOnClick(it, i) : '';
    const isSelected = options.getIsSelected ? options.getIsSelected(it, i) : false;
    const legAttrs = onClick
      ? ` role="button" tabindex="0" onclick="${escAttr(onClick)}" onkeydown="${escAttr(`if(event.key==='Enter'||event.key===' '){event.preventDefault();${onClick}}`)}" title="${escAttr(options.clickTitle || 'Filtrar')}" data-tooltip="${escAttr(options.tooltipPrefix ? options.tooltipPrefix + it.name : 'Filtrar por ' + it.name)}"`
      : '';
    return `<div class="leg-item ${onClick ? 'leg-item-action' : ''} ${isSelected ? 'chart-selected' : ''}"${legAttrs} style="display:flex;align-items:center;gap:8px;font-size:12px"><div class="leg-dot" style="background:${COLORS[i%COLORS.length]};width:10px;height:10px;border-radius:50%;flex-shrink:0"></div><span title="${escAttr(it.name)}" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text-strong);font-weight:600">${escHtml(it.name)}</span><span class="leg-pct" style="color:var(--text);font-family:var(--font-mono);font-weight:600">${pct}%</span></div>`;
  }).join('');
  attachChartTooltips(svg.parentElement || svg);
}

function renderEvoChart(containerId, dataByDir, months, opts){
  const el=document.getElementById(containerId);
  if(!el) return;
  const options = opts || {};
  const monthKeys = (months && months.length)
    ? months
    : [...new Set(Object.values(dataByDir || {}).flatMap(row => Object.keys(row || {})).filter(Boolean))].sort();
  if(!monthKeys.length){
    el.innerHTML = `<div style="font-size:11px;color:var(--text2)">Sin datos mensuales para mostrar.</div>`;
    return;
  }
  const dirs=options.dirOrder && options.dirOrder.length
    ? options.dirOrder.filter(d => Object.prototype.hasOwnProperty.call(dataByDir, d))
    : [...new Set(Object.keys(dataByDir))];
  const nDirs=dirs.length;
  const allVals=dirs.flatMap(d=>monthKeys.map(m=>dataByDir[d][m]||0));
  const maxVal=Math.max(...allVals,1);
  const positiveVals = allVals.filter(v => v > 0).sort((a,b)=>b-a);
  const secondVal = positiveVals.find(v => v < maxVal * 0.98) || positiveVals[1] || 0;
  const hasOutlier = secondVal > 0 && maxVal / secondVal >= 5;
  const displayMax = hasOutlier ? Math.max(secondVal * 1.35, 1) : maxVal;
  const W=520,H=218,padL=62,padR=12,padT=hasOutlier?32:26,padB=38;
  const gW=W-padL-padR, gH=H-padT-padB;
  const grpW=gW/monthKeys.length;
  const barW=Math.min(18, (grpW-8)/Math.max(nDirs,1));
  const gap=2;
  const legendHtml=`<div class="evo-legend" aria-label="Directores">${dirs.map((d,di)=>{
    const c=COLORS[di%COLORS.length];
    return `<div class="evo-legend-item" title="${escAttr(d)}"><span class="evo-legend-dot" style="background:${c}"></span><span class="evo-legend-name">${escHtml(d)}</span></div>`;
  }).join('')}</div>`;

  let svg=`<svg viewBox="0 0 ${W} ${H}" class="evo-chart-svg" aria-hidden="true">`;
  svg+=`<defs>${dirs.map((d,di)=>`<linearGradient id="bg${di}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${COLORS[di%COLORS.length]}" stop-opacity=".95"/><stop offset="100%" stop-color="${COLORS[di%COLORS.length]}" stop-opacity=".4"/></linearGradient>`).join('')}</defs>`;

  // Grid lines
  [0,.25,.5,.75,1].forEach(t=>{
    const y=padT+gH*(1-t);
    svg+=`<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="var(--border)" stroke-width="${t===0?1.5:.7}"/>`;
    if(t>0) svg+=`<text x="${padL-5}" y="${y+3.5}" text-anchor="end" font-size="8.5" font-weight="400" fill="var(--text3)" font-family="IBM Plex Mono,monospace">${abr(displayMax*t)}</text>`;
  });

  // Bars
  monthKeys.forEach((m,mi)=>{
    const grpCenter=padL+(mi+0.5)*grpW;
    const totalBarW=(barW+gap)*nDirs-gap;
    dirs.forEach((d,di)=>{
      const v=dataByDir[d][m]||0;
      const clipped = hasOutlier && v > displayMax;
      const displayVal = Math.min(v, displayMax);
      const bh=Math.max(displayVal/displayMax*gH, v>0?2:0);
      const x=grpCenter - totalBarW/2 + di*(barW+gap);
      const y=padT+gH-bh;
      svg+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW}" height="${bh.toFixed(1)}" rx="2" fill="url(#bg${di})" data-tooltip="${escAttr(dirs[di]+': '+abr(v))}"></rect>`;
      if(clipped) {
        const c = COLORS[di%COLORS.length];
        const cx = x + barW/2;
        svg+=`<path d="M${(x+2).toFixed(1)},${(padT+8).toFixed(1)} L${(x+barW-2).toFixed(1)},${(padT+3).toFixed(1)} M${(x+2).toFixed(1)},${(padT+14).toFixed(1)} L${(x+barW-2).toFixed(1)},${(padT+9).toFixed(1)}" stroke="${c}" stroke-width="1.5" stroke-linecap="round" opacity=".95"></path>`;
        svg+=`<text x="${cx.toFixed(1)}" y="${(padT-9).toFixed(1)}" text-anchor="middle" font-size="7.5" font-weight="700" fill="${c}" font-family="IBM Plex Mono,monospace">${abr(v)}</text>`;
      }
    });
    // Month label
    svg+=`<text x="${grpCenter.toFixed(1)}" y="${H-8}" text-anchor="middle" font-size="10" fill="var(--text3)" font-family="IBM Plex Sans,sans-serif" font-weight="400">${getMonthShortLabel(m)}</text>`;
  });

  svg+='</svg>';
  el.innerHTML=`<div class="evo-chart-wrap">${legendHtml}${svg}</div>`;
  attachChartTooltips(el);
}

/* ══════════════════════════════════════
   GERENCIA GENERAL
══════════════════════════════════════ */
function renderGerencia(){
  const BASE_DATA = getVisibleData();
  if(!BASE_DATA.length) return;
  const ALL_DATA = applyGerenciaCrossfilters(BASE_DATA);
  renderGerenciaCrossfilterBar(BASE_DATA.length, ALL_DATA.length);
  const trm=getTRM();
  const today=todayStr();
  
  // Hoy strip
  const hoy=ALL_DATA.filter(r=>{
    const f=r['FECHA DIA/MES/AÑO'];
    if(!f) return false;
    const d=f instanceof Date?f.toISOString().substring(0,10):(typeof f==='number'?new Date(Math.round((f-25569)*86400*1000)).toISOString().substring(0,10):String(f)).substring(0,10);
    return d===today;
  });
  document.getElementById('hoy-fecha').textContent=new Date().toLocaleDateString('es-CO',{weekday:'long',year:'numeric',month:'long',day:'numeric'});
  const hoyCOP=hoy.filter(r=>(r['MONEDA 2']||'').trim()==='COP').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const hoyUSD=hoy.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  document.getElementById('hoy-cop').textContent=fmtCOP(hoyCOP);
  document.getElementById('hoy-usd').textContent=fmtCOP(hoyUSD*trm);
  document.getElementById('hoy-count').textContent=hoy.length;
  document.getElementById('hoy-trm').textContent='$ '+fmtTRMDisplay(trm);
  
  // Total COP (all)
  const totalCOP=ALL_DATA.reduce((s,r)=>s+toCOP(r),0);
  const totalUSD=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const totalCOPonly=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='COP').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const ganadas=ALL_DATA.filter(r=>r['ESTADO']==='GANADA');
  const totalGanada=ganadas.reduce((s,r)=>s+toCOP(r),0);
  
  document.getElementById('kpi-gerencia').innerHTML=`
    <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total Forecast COP</div>
      <div class="kpi-val">${abr(totalCOP)}</div>
      <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      <span class="kpi-badge cop">COP</span>
    </div>
    <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total USD</div>
      ${kpiUSD(totalUSD)}
      <div class="kpi-sub">Liq: ${abr(totalUSD*trm)}</div>
      <span class="kpi-badge usd">USD</span>
    </div>
    <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
      <div class="kpi-label">Ventas Ganadas</div>
      <div class="kpi-val">${abr(totalGanada)}</div>
      <div class="kpi-sub">${ganadas.length} negocios</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
      <div class="kpi-label">Negocios Totales</div>
      <div class="kpi-val">${ALL_DATA.length}</div>
      <div class="kpi-sub">${[...new Set(ALL_DATA.map(r=>r['COMERCIAL']))].length} ejecutivos</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
      <div class="kpi-label">TRM Día</div>
      <div class="kpi-val">$ ${fmtTRMDisplay(trm)}</div>
      <div class="kpi-sub">COP por USD</div>
    </div>
  `;
  
  // Bar directores — todos los de LOADED_FILES_BY_DIR aunque no tengan datos
  const dirsFromData=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))];
  const dirsFromFiles=Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim()).filter(Boolean);
  const dirsUniq=[...new Set([...dirsFromData,...dirsFromFiles])].sort();
  const dirData=dirsUniq.map((d)=>({
    name:d, val:ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d).reduce((s,r)=>s+toCOP(r),0)
  })).sort((a,b)=>b.val-a.val);
  renderBars('bar-directores',dirData,COLORS,null,{
    getOnClick: item => jsCall('setGerenciaCrossfilter', 'director', item.name),
    getIsSelected: item => GERENCIA_CROSSFILTERS.director === item.name,
    clickTitle: 'Filtrar por director',
    tooltipPrefix: 'Filtrar director: '
  });
  
  // Evo por director mensual
  const monthsForEvo = getForecastMonths(ALL_DATA);
  const dirsForEvo=dirData.map(item => item.name);
  const evoByDir={};
  dirsForEvo.forEach(d=>{
    evoByDir[d]={};
    monthsForEvo.forEach(m=>{
      evoByDir[d][m]=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d&&getMonth(getRowDateValue(r))===m).reduce((s,r)=>s+toCOP(r),0);
    });
  });
  renderEvoChart('evo-dir-chart',evoByDir, monthsForEvo, { dirOrder: dirsForEvo });
  
  // Bar ejecutivos
  const execs=[...new Set(ALL_DATA.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  const ejData=execs.map(e=>({name:e.split(' ')[0],filterValue:e,val:ALL_DATA.filter(r=>r['COMERCIAL']===e).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderBars('bar-ejecutivos',ejData.slice(0, TOP_BAR_LIMIT),COLORS,null,{
    getOnClick: item => jsCall('setGerenciaCrossfilter', 'ejecutivo', item.filterValue || item.name),
    getIsSelected: item => namesMatch(item.filterValue || item.name, GERENCIA_CROSSFILTERS.ejecutivo),
    clickTitle: 'Filtrar por ejecutivo',
    tooltipPrefix: 'Filtrar ejecutivo: '
  });
  
  // Donuts
  const estados=['GANADA','PENDIENTE','PERDIDA','APLAZADO'];
  const estData=estados.map(e=>({name:e,val:ALL_DATA.filter(r=>r['ESTADO']===e).reduce((s,r)=>s+toCOP(r),0)}));
  renderDonut('donut-estado','leg-estado',estData,{
    hasActive: Boolean(GERENCIA_CROSSFILTERS.estado),
    getOnClick: item => jsCall('setGerenciaCrossfilter', 'estado', item.name),
    getIsSelected: item => GERENCIA_CROSSFILTERS.estado === item.name,
    clickTitle: 'Filtrar por estado',
    tooltipPrefix: 'Filtrar estado: '
  });
  
  const linData=buildLineValueData(ALL_DATA);
  renderDonut('donut-linea','leg-linea',linData,{
    hasActive: Boolean(GERENCIA_CROSSFILTERS.linea),
    getOnClick: item => jsCall('setGerenciaCrossfilter', 'linea', item.name),
    getIsSelected: item => GERENCIA_CROSSFILTERS.linea === item.name,
    clickTitle: 'Filtrar por linea',
    tooltipPrefix: 'Filtrar linea: '
  });

  // Tablas de estados
  renderGerenciaEstadoTables(ALL_DATA);
}

function renderGerenciaEstadoTables(data) {
  const el = document.getElementById('gerencia-estado-tables');
  if(!el) return;
  const estados = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'];
  const colores = {GANADA:'#0DBF82',PENDIENTE:'#F0A020',PERDIDA:'#2D4FD6',APLAZADO:'#E84040'};

  el.innerHTML = estados.map(estado => {
    const rows = data
      .filter(r => cleanDisplayText(r['ESTADO'], '').toUpperCase() === estado)
      .sort((a,b) => toCOP(b) - toCOP(a));
    const total = rows.reduce((s,r) => s + toCOP(r), 0);
    const visible = GERENCIA_ESTADO_LIMITS[estado] || 30;

    if(!rows.length) return `
      <div class="estado-card estado-card-empty" style="background:var(--card);border:1px solid var(--border);border-left:3px solid ${colores[estado]};border-radius:12px;padding:16px">
        <div style="font-family:var(--font-display);font-size:10px;font-weight:700;letter-spacing:1px;color:${colores[estado]};margin-bottom:6px">${estado}</div>
        <div style="font-size:11px;color:var(--text3)">Sin registros</div>
      </div>`;

    const rows_html = rows.slice(0, visible).map(r => {
      const cliente = cleanDisplayText(getRowClientName(r), 'Sin cliente');
      const comercialFull = cleanDisplayText(r['COMERCIAL'], 'Sin ejecutivo');
      const comercial = comercialFull.split(' ')[0] || comercialFull;
      const valor = toCOP(r);
      const linea = cleanDisplayText(getRowLineName(r), 'Sin linea');
      return `
        <tr class="table-row-action" style="border-top:1px solid var(--border)" onclick="${escAttr(jsCall('openNegocioDetailById', r.__RID, 'gerencia'))}" title="Abrir detalle del negocio">
          <td style="padding:5px 8px;font-size:10px;color:var(--text);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escAttr(cliente)}">${escHtml(cliente)}</td>
          <td style="padding:5px 8px;font-size:10px;color:var(--text3);font-weight:600;white-space:nowrap">${escHtml(comercial)}</td>
          <td style="padding:5px 8px;font-size:10px;color:var(--text3);font-weight:500;white-space:nowrap;max-width:80px;overflow:hidden;text-overflow:ellipsis" title="${escAttr(linea)}">${escHtml(linea)}</td>
          <td style="padding:5px 8px;font-size:10px;color:${colores[estado]};text-align:right;font-family:var(--font-mono);font-weight:600;white-space:nowrap">${abr(valor)}</td>
        </tr>`;
    }).join('');

    const remaining = Math.max(rows.length - visible, 0);
    const more = remaining > 0 ? `
      <div class="table-more-wrap">
        <button type="button" class="table-more-btn" onclick="showMoreEstadoRows('${estado}')">Ver mas (${remaining})</button>
      </div>` : '';

    return `
      <div class="estado-card" style="background:var(--card);border:1px solid var(--border);border-left:3px solid ${colores[estado]};border-radius:12px;overflow:hidden">
        <div class="estado-card-head" style="padding:10px 14px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--border)">
          <div>
            <span style="font-family:var(--font-display);font-size:10px;font-weight:700;letter-spacing:1px;color:${colores[estado]}">${estado}</span>
            <span style="font-size:9px;color:var(--text3);margin-left:8px;font-family:var(--font-body)">${rows.length} negocio${rows.length!==1?'s':''}</span>
          </div>
          <span style="font-family:var(--font-mono);font-size:13px;font-weight:700;color:var(--text)">${abr(total)}</span>
        </div>
        <div class="estado-card-body" style="overflow-y:auto;max-height:280px" data-scroll-key="estado:${estado}">
          <table class="estado-mini-table" style="width:100%;border-collapse:collapse">
            <thead style="position:sticky;top:0;z-index:1;background:var(--bg2)">
              <tr>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:left">EMPRESA</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:left">EJECUTIVO</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:left">LÍNEA</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:right">VALOR</th>
              </tr>
            </thead>
            <tbody>${rows_html}</tbody>
            <tfoot style="background:var(--bg2);border-top:1px solid var(--border)">
              <tr>
                <td colspan="3" style="padding:6px 8px;font-size:9px;font-family:var(--font-display);color:var(--text);font-weight:700">TOTAL ${rows.length} NEGOCIOS</td>
                <td style="padding:6px 8px;font-size:11px;text-align:right;font-family:var(--font-mono);color:${colores[estado]};font-weight:700">${abr(total)}</td>
              </tr>
            </tfoot>
          </table>
        </div>${more}
      </div>`;
  }).join('');
}

function excelArgb(hex){
  return 'FF' + String(hex || '#000000').replace('#','').toUpperCase().padStart(6,'0').slice(0,6);
}

function excelColumnLetter(index){
  let n = Number(index) || 1;
  let letter = '';
  while(n > 0){
    const mod = (n - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    n = Math.floor((n - mod) / 26);
  }
  return letter;
}

function excelSafeSheetName(name){
  const clean = cleanDisplayText(name, 'Hoja')
    .replace(/[\[\]\*\/\\\?:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return (clean || 'Hoja').substring(0, 31);
}

function excelSafeText(value, fallback){
  const text = cleanDisplayText(value, fallback || '');
  if(!text) return '';
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function excelDateValue(value){
  const formatted = formatDateValue(value);
  if(/^\d{4}-\d{2}-\d{2}$/.test(formatted)){
    const parts = formatted.split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  return formatted || '';
}

function excelMoneyFormat(moneda, withDecimals){
  const currency = cleanDisplayText(moneda, 'COP').trim().toUpperCase();
  if(currency === 'USD') return withDecimals ? '[$USD] #,##0.##' : '[$USD] #,##0';
  return withDecimals ? '[$$-es-CO] #,##0.##' : '[$$-es-CO] #,##0';
}

function getGerenciaEstadoExportData(){
  return applyGerenciaCrossfilters(getVisibleData())
    .filter(row => GERENCIA_ESTADOS.includes(cleanDisplayText(row['ESTADO'], '').toUpperCase()))
    .sort((a,b) => {
      const stateA = GERENCIA_ESTADOS.indexOf(cleanDisplayText(a['ESTADO'], '').toUpperCase());
      const stateB = GERENCIA_ESTADOS.indexOf(cleanDisplayText(b['ESTADO'], '').toUpperCase());
      if(stateA !== stateB) return stateA - stateB;
      return toCOP(b) - toCOP(a);
    });
}

function getGerenciaEstadoFilterText(){
  const labels = {
    estado: 'Estado',
    director: 'Director',
    ejecutivo: 'Ejecutivo',
    linea: 'Linea'
  };
  const filters = Object.entries(GERENCIA_CROSSFILTERS || {})
    .filter(([,value]) => value)
    .map(([key,value]) => `${labels[key] || key}: ${value}`);
  const role = CURRENT_USER ? getRoleLabel(CURRENT_USER.role) : 'Sin usuario';
  return [`Vista: ${role}`, ...filters].join(' | ') || 'Sin filtros';
}

function getBogotaTimestampForFile(){
  const bogota = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  const pad = n => String(n).padStart(2,'0');
  return `${bogota.getFullYear()}${pad(bogota.getMonth()+1)}${pad(bogota.getDate())}_${pad(bogota.getHours())}${pad(bogota.getMinutes())}`;
}

function getBogotaTimestampLabel(){
  return new Date().toLocaleString('es-CO', {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function getGerenciaExportRowValues(row){
  const moneda = cleanDisplayText(row['MONEDA 2'], 'COP').trim().toUpperCase() || 'COP';
  const valorOriginal = parseMonto(row['MONTO VENTA CLIENTE']) || 0;
  const costoOriginal = parseMonto(firstFilled(row, ['COSTO NEGOCIO','COSTO'])) || 0;
  const costoCop = moneda === 'USD' ? costoOriginal * getTRM() : costoOriginal;
  const utilidad = getUtilidad(row);
  const utilidadCop = utilidad.moneda === 'USD' ? utilidad.valor * getTRM() : utilidad.valor;
  const margen = getMarginRatio(row['MARGEN']);
  const numero = firstFilled(row, ['NUMERO DE COTIZACION']) || getRowPartNumber(row);

  return [
    excelSafeText(cleanDisplayText(row['ESTADO'], 'Sin estado').toUpperCase()),
    excelSafeText(getRowClientName(row), 'Sin cliente'),
    excelSafeText(row['DIRECTOR'], 'Sin director'),
    excelSafeText(row['COMERCIAL'], 'Sin ejecutivo'),
    excelSafeText(getRowLineName(row), 'Sin linea'),
    excelSafeText(getRowBrandName(row), 'Sin marca'),
    excelSafeText(getRowProductName(row), 'Sin producto'),
    getRowQuantityValue(row, null),
    excelSafeText(numero, 'Sin numero'),
    excelDateValue(getRowDateValue(row)),
    excelSafeText(moneda, 'COP'),
    valorOriginal,
    costoOriginal,
    getTRM(),
    toCOP(row),
    costoCop,
    margen,
    utilidadCop || 0,
    excelSafeText(row.__SOURCE_FILE, ''),
    excelSafeText(row.__SOURCE_SHEET, ''),
    excelSafeText(getRowObservation(row), '')
  ];
}

function getGerenciaEstadoSummaryRows(rows){
  const grandTotal = rows.reduce((sum,row) => sum + toCOP(row), 0);
  return GERENCIA_ESTADOS.map(estado => {
    const stateRows = rows.filter(row => cleanDisplayText(row['ESTADO'], '').toUpperCase() === estado);
    const totalCOP = stateRows.reduce((sum,row) => sum + toCOP(row), 0);
    const totalUSD = stateRows
      .filter(row => cleanDisplayText(row['MONEDA 2'], 'COP').trim().toUpperCase() === 'USD')
      .reduce((sum,row) => sum + (parseMonto(row['MONTO VENTA CLIENTE']) || 0), 0);
    const utilidadCOP = stateRows.reduce((sum,row) => {
      const utilidad = getUtilidad(row);
      return sum + (utilidad.moneda === 'USD' ? utilidad.valor * getTRM() : utilidad.valor);
    }, 0);
    return [estado, stateRows.length, totalCOP, totalUSD, utilidadCOP, grandTotal ? totalCOP / grandTotal : 0];
  });
}

function getGerenciaDirectorSummaryRows(rows){
  const directors = [...new Set(rows.map(row => cleanDisplayText(row['DIRECTOR'], 'Sin director')))].sort();
  return directors.map(director => {
    const directorRows = rows.filter(row => cleanDisplayText(row['DIRECTOR'], 'Sin director') === director);
    const totalCOP = directorRows.reduce((sum,row) => sum + toCOP(row), 0);
    const totalUSD = directorRows
      .filter(row => cleanDisplayText(row['MONEDA 2'], 'COP').trim().toUpperCase() === 'USD')
      .reduce((sum,row) => sum + (parseMonto(row['MONTO VENTA CLIENTE']) || 0), 0);
    const counts = GERENCIA_ESTADOS.reduce((acc, estado) => {
      acc[estado] = directorRows.filter(row => cleanDisplayText(row['ESTADO'], '').toUpperCase() === estado).length;
      return acc;
    }, {});
    return [
      director,
      directorRows.length,
      totalCOP,
      totalUSD,
      counts.GANADA,
      counts.PENDIENTE,
      counts.PERDIDA,
      counts.APLAZADO
    ];
  }).sort((a,b) => b[2] - a[2]);
}

function styleExcelTitle(ws, title, subtitle, filterText, lastCol){
  const last = excelColumnLetter(lastCol);
  ws.mergeCells(`A1:${last}1`);
  ws.mergeCells(`A2:${last}2`);
  ws.mergeCells(`A3:${last}3`);
  ws.getRow(1).height = 28;
  ws.getRow(2).height = 22;
  ws.getRow(3).height = 22;
  ws.getCell('A1').value = title;
  ws.getCell('A1').font = { name:'Aptos Display', size:16, bold:true, color:{ argb:'FFFFFFFF' } };
  ws.getCell('A1').fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#1B2B8C') } };
  ws.getCell('A1').alignment = { vertical:'middle', horizontal:'center' };
  ws.getCell('A2').value = subtitle;
  ws.getCell('A2').font = { name:'Aptos', size:10, bold:true, color:{ argb: excelArgb('#1A2B6B') } };
  ws.getCell('A2').fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#EEF3FF') } };
  ws.getCell('A2').alignment = { vertical:'middle', horizontal:'center' };
  ws.getCell('A3').value = `Filtros: ${filterText}`;
  ws.getCell('A3').font = { name:'Aptos', size:9, italic:true, color:{ argb: excelArgb('#677592') } };
  ws.getCell('A3').fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
  ws.getCell('A3').alignment = { vertical:'middle', horizontal:'center' };
}

function styleExcelSectionLabel(ws, rowNumber, text, lastCol){
  const last = excelColumnLetter(lastCol);
  ws.mergeCells(`A${rowNumber}:${last}${rowNumber}`);
  const cell = ws.getCell(`A${rowNumber}`);
  cell.value = text;
  cell.font = { name:'Aptos Display', size:11, bold:true, color:{ argb: excelArgb('#1A2B6B') } };
  cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#EEF3FF') } };
  cell.alignment = { vertical:'middle', horizontal:'left' };
  ws.getRow(rowNumber).height = 22;
}

function styleExcelKpis(ws, rows){
  const totalCOP = rows.reduce((sum,row) => sum + toCOP(row), 0);
  const totalUSD = rows
    .filter(row => cleanDisplayText(row['MONEDA 2'], 'COP').trim().toUpperCase() === 'USD')
    .reduce((sum,row) => sum + (parseMonto(row['MONTO VENTA CLIENTE']) || 0), 0);
  const cards = [
    ['Negocios', rows.length, '#2D4FD6', '#,##0'],
    ['Total COP', totalCOP, '#0DBF82', excelMoneyFormat('COP', false)],
    ['Total USD', totalUSD, '#2D4FD6', excelMoneyFormat('USD', false)],
    ['TRM aplicada', getTRM(), '#2ABFDF', excelMoneyFormat('COP', true)]
  ];
  cards.forEach((card, idx) => {
    const startCol = idx * 2 + 1;
    const endCol = startCol + 1;
    ws.mergeCells(5, startCol, 5, endCol);
    ws.mergeCells(6, startCol, 6, endCol);
    const label = ws.getCell(5, startCol);
    const value = ws.getCell(6, startCol);
    label.value = card[0];
    value.value = card[1];
    value.numFmt = card[3];
    [label, value].forEach(cell => {
      cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
      cell.border = {
        top:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
        left:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
        bottom:{style:'thin', color:{argb:excelArgb('#D7E0F0')}},
        right:{style:'thin', color:{argb:excelArgb('#D7E0F0')}}
      };
      cell.alignment = { vertical:'middle', horizontal:'center' };
    });
    label.font = { name:'Aptos', size:9, bold:true, color:{ argb: excelArgb('#677592') } };
    value.font = { name:'Aptos Display', size:13, bold:true, color:{ argb: excelArgb(card[2]) } };
  });
  ws.getRow(5).height = 20;
  ws.getRow(6).height = 24;
}

function addExcelTable(ws, tableConfig){
  const table = ws.addTable({
    name: tableConfig.name,
    ref: tableConfig.ref,
    headerRow: true,
    totalsRow: tableConfig.totalsRow !== false,
    style: {
      theme: tableConfig.theme || 'TableStyleMedium2',
      showRowStripes: true
    },
    columns: tableConfig.columns,
    rows: tableConfig.rows
  });
  if(table && typeof table.commit === 'function') table.commit();
}

function styleExcelTable(ws, startRow, rowCount, colCount, options){
  const opts = options || {};
  const headerFill = excelArgb(opts.headerColor || '#1B2B8C');
  const borderColor = excelArgb('#D7E0F0');
  const header = ws.getRow(startRow);
  header.height = 22;
  header.eachCell((cell) => {
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: headerFill } };
    cell.font = { name:'Aptos', size:9, bold:true, color:{ argb:'FFFFFFFF' } };
    cell.alignment = { vertical:'middle', horizontal:'center', wrapText:true };
    cell.border = {
      top:{style:'thin', color:{argb:borderColor}},
      left:{style:'thin', color:{argb:borderColor}},
      bottom:{style:'thin', color:{argb:borderColor}},
      right:{style:'thin', color:{argb:borderColor}}
    };
  });

  const lastRow = startRow + rowCount + 1;
  for(let rowNumber = startRow + 1; rowNumber <= lastRow; rowNumber++){
    const row = ws.getRow(rowNumber);
    row.eachCell((cell) => {
      cell.font = { name:'Aptos', size:9, color:{ argb: excelArgb('#172033') } };
      cell.alignment = { vertical:'top', horizontal:'left', wrapText:false };
      cell.border = {
        top:{style:'thin', color:{argb:borderColor}},
        left:{style:'thin', color:{argb:borderColor}},
        bottom:{style:'thin', color:{argb:borderColor}},
        right:{style:'thin', color:{argb:borderColor}}
      };
    });
  }

  if(opts.statusColumn){
    for(let rowNumber = startRow + 1; rowNumber <= startRow + rowCount; rowNumber++){
      const statusCell = ws.getCell(rowNumber, opts.statusColumn);
      const color = GERENCIA_ESTADO_COLORS[String(statusCell.value || '').toUpperCase()];
      if(color){
        statusCell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb(color) } };
        statusCell.font = { name:'Aptos', size:9, bold:true, color:{ argb:'FFFFFFFF' } };
        statusCell.alignment = { vertical:'middle', horizontal:'center' };
      }
    }
  }

  if(opts.wrapColumns){
    for(let rowNumber = startRow + 1; rowNumber <= startRow + rowCount; rowNumber++){
      opts.wrapColumns.forEach(col => {
        ws.getCell(rowNumber, col).alignment = { vertical:'top', horizontal:'left', wrapText:true };
      });
    }
  }

  const total = ws.getRow(lastRow);
  total.height = 22;
  for(let col = 1; col <= colCount; col++){
    const cell = ws.getCell(lastRow, col);
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#EEF3FF') } };
    cell.font = { name:'Aptos', size:9, bold:true, color:{ argb: excelArgb('#1A2B6B') } };
    cell.alignment = { vertical:'middle', horizontal: opts.rightColumns && opts.rightColumns.includes(col) ? 'right' : 'left' };
  }
}

function formatDetailExcelColumns(ws, startRow, sourceRows){
  const rowCount = sourceRows.length;
  const totalRow = startRow + rowCount + 1;
  for(let rowNumber = startRow + 1; rowNumber <= totalRow; rowNumber++){
    const sourceRow = sourceRows[rowNumber - startRow - 1] || {};
    const moneda = cleanDisplayText(sourceRow['MONEDA 2'], 'COP').trim().toUpperCase();
    ws.getCell(rowNumber, 8).numFmt = '#,##0';
    ws.getCell(rowNumber, 10).numFmt = 'yyyy-mm-dd';
    ws.getCell(rowNumber, 12).numFmt = excelMoneyFormat(moneda, false);
    ws.getCell(rowNumber, 13).numFmt = excelMoneyFormat(moneda, false);
    ws.getCell(rowNumber, 14).numFmt = excelMoneyFormat('COP', true);
    ws.getCell(rowNumber, 15).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 16).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 17).numFmt = '0.##%';
    ws.getCell(rowNumber, 18).numFmt = excelMoneyFormat('COP', false);
    [8,12,13,14,15,16,17,18].forEach(col => {
      ws.getCell(rowNumber, col).alignment = { vertical:'top', horizontal:'right' };
    });
  }
}

function addDetailExcelTable(ws, rows, tableName, startRow){
  const columns = [
    { name:'Estado', totalsRowLabel:'Total' },
    { name:'Empresa' },
    { name:'Director' },
    { name:'Ejecutivo' },
    { name:'Linea' },
    { name:'Marca' },
    { name:'Producto' },
    { name:'Cantidad', totalsRowFunction:'sum' },
    { name:'Cotizacion / Parte' },
    { name:'Fecha' },
    { name:'Moneda' },
    { name:'Valor original' },
    { name:'Costo original' },
    { name:'TRM aplicada' },
    { name:'Valor total COP', totalsRowFunction:'sum' },
    { name:'Costo COP liq.', totalsRowFunction:'sum' },
    { name:'Margen' },
    { name:'Utilidad COP liq.', totalsRowFunction:'sum' },
    { name:'Archivo fuente' },
    { name:'Hoja fuente' },
    { name:'Observaciones' }
  ];
  const widths = [14,28,24,24,20,18,32,11,20,13,10,15,15,14,17,16,11,17,24,16,42];
  widths.forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });
  addExcelTable(ws, {
    name: tableName,
    ref: `A${startRow}`,
    columns,
    rows: rows.map(getGerenciaExportRowValues),
    theme: 'TableStyleMedium2'
  });
  styleExcelTable(ws, startRow, rows.length, columns.length, {
    statusColumn: 1,
    wrapColumns: [2,7,19,21],
    rightColumns: [8,12,13,14,15,16,17,18]
  });
  formatDetailExcelColumns(ws, startRow, rows);
}

function setupExcelWorksheet(ws, lastCol, frozenRows){
  ws.properties.defaultRowHeight = 18;
  ws.views = [{ state:'frozen', ySplit: frozenRows || 0 }];
  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left:0.25, right:0.25, top:0.4, bottom:0.4, header:0.2, footer:0.2 }
  };
  ws.headerFooter.oddFooter = '&LForecast 2026&C&P de &N&RProvexpress';
  for(let col = 1; col <= lastCol; col++){
    ws.getColumn(col).font = { name:'Aptos', size:9 };
  }
}

function addResumenExcelSheet(workbook, rows){
  const ws = workbook.addWorksheet('Resumen');
  setupExcelWorksheet(ws, 8, 9);
  styleExcelTitle(
    ws,
    'Forecast 2026 - Detalle negocios por estado',
    `Generado: ${getBogotaTimestampLabel()} | Registros exportados: ${rows.length}`,
    getGerenciaEstadoFilterText(),
    8
  );
  styleExcelKpis(ws, rows);

  styleExcelSectionLabel(ws, 8, 'Resumen por estado', 8);
  const summaryColumns = [
    { name:'Estado', totalsRowLabel:'Total' },
    { name:'Negocios', totalsRowFunction:'sum' },
    { name:'Total COP', totalsRowFunction:'sum' },
    { name:'Total USD original', totalsRowFunction:'sum' },
    { name:'Utilidad COP liq.', totalsRowFunction:'sum' },
    { name:'Participacion', totalsRowFunction:'sum' }
  ];
  const summaryRows = getGerenciaEstadoSummaryRows(rows);
  [16,12,18,18,18,14].forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });
  addExcelTable(ws, {
    name: 'ResumenEstados',
    ref: 'A9',
    columns: summaryColumns,
    rows: summaryRows,
    theme: 'TableStyleMedium4'
  });
  styleExcelTable(ws, 9, summaryRows.length, summaryColumns.length, {
    statusColumn: 1,
    rightColumns: [2,3,4,5,6]
  });
  for(let rowNumber = 10; rowNumber <= 14; rowNumber++){
    ws.getCell(rowNumber, 2).numFmt = '#,##0';
    ws.getCell(rowNumber, 3).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 4).numFmt = excelMoneyFormat('USD', false);
    ws.getCell(rowNumber, 5).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 6).numFmt = '0.##%';
  }

  const directorStart = 17;
  styleExcelSectionLabel(ws, directorStart - 1, 'Resumen por director', 8);
  const directorColumns = [
    { name:'Director', totalsRowLabel:'Total' },
    { name:'Negocios', totalsRowFunction:'sum' },
    { name:'Total COP', totalsRowFunction:'sum' },
    { name:'Total USD original', totalsRowFunction:'sum' },
    { name:'Ganada', totalsRowFunction:'sum' },
    { name:'Pendiente', totalsRowFunction:'sum' },
    { name:'Perdida', totalsRowFunction:'sum' },
    { name:'Aplazado', totalsRowFunction:'sum' }
  ];
  const directorRows = getGerenciaDirectorSummaryRows(rows);
  addExcelTable(ws, {
    name: 'ResumenDirectores',
    ref: `A${directorStart}`,
    columns: directorColumns,
    rows: directorRows,
    theme: 'TableStyleMedium9'
  });
  styleExcelTable(ws, directorStart, directorRows.length, directorColumns.length, {
    rightColumns: [2,3,4,5,6,7,8]
  });
  for(let rowNumber = directorStart + 1; rowNumber <= directorStart + directorRows.length + 1; rowNumber++){
    ws.getCell(rowNumber, 2).numFmt = '#,##0';
    ws.getCell(rowNumber, 3).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 4).numFmt = excelMoneyFormat('USD', false);
    [5,6,7,8].forEach(col => { ws.getCell(rowNumber, col).numFmt = '#,##0'; });
  }
}

function addEstadoExcelSheet(workbook, estado, rows){
  const ws = workbook.addWorksheet(excelSafeSheetName(estado));
  setupExcelWorksheet(ws, GERENCIA_EXCEL_DETAIL_COLS, 6);
  styleExcelTitle(
    ws,
    `Detalle estado ${estado}`,
    `Generado: ${getBogotaTimestampLabel()} | Registros: ${rows.length}`,
    getGerenciaEstadoFilterText(),
    GERENCIA_EXCEL_DETAIL_COLS
  );
  if(!rows.length){
    ws.mergeCells(`A6:${excelColumnLetter(GERENCIA_EXCEL_DETAIL_COLS)}6`);
    const cell = ws.getCell('A6');
    cell.value = 'Sin registros para este estado con los filtros actuales.';
    cell.font = { name:'Aptos', size:11, italic:true, color:{ argb: excelArgb('#677592') } };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
    cell.alignment = { vertical:'middle', horizontal:'center' };
    ws.getRow(6).height = 24;
    return;
  }
  addDetailExcelTable(ws, rows, `Detalle${estado}`, 6);
}

function addDetalleGeneralExcelSheet(workbook, rows){
  const ws = workbook.addWorksheet('Detalle General');
  setupExcelWorksheet(ws, GERENCIA_EXCEL_DETAIL_COLS, 6);
  styleExcelTitle(
    ws,
    'Detalle general por estado',
    `Generado: ${getBogotaTimestampLabel()} | Registros: ${rows.length}`,
    getGerenciaEstadoFilterText(),
    GERENCIA_EXCEL_DETAIL_COLS
  );
  addDetailExcelTable(ws, rows, 'DetalleGeneralEstados', 6);
}

function getGerenciaComiteRows(rows){
  return (rows || [])
    .filter(row => toCOP(row) > GERENCIA_COMITE_THRESHOLD_COP)
    .sort((a,b) => toCOP(b) - toCOP(a));
}

function addComiteGerenciaExcelSheet(workbook, rows){
  const comiteRows = getGerenciaComiteRows(rows);
  const ws = workbook.addWorksheet('Someter comite gerencia general');
  setupExcelWorksheet(ws, GERENCIA_EXCEL_DETAIL_COLS, 6);
  styleExcelTitle(
    ws,
    'Someter a comite de gerencia general',
    `Negocios mayores a ${fmtCOP(GERENCIA_COMITE_THRESHOLD_COP)} | Registros: ${comiteRows.length}`,
    getGerenciaEstadoFilterText(),
    GERENCIA_EXCEL_DETAIL_COLS
  );
  if(!comiteRows.length){
    ws.mergeCells(`A6:${excelColumnLetter(GERENCIA_EXCEL_DETAIL_COLS)}6`);
    const cell = ws.getCell('A6');
    cell.value = 'No hay negocios mayores a $1.000 millones COP con los filtros actuales.';
    cell.font = { name:'Aptos', size:11, italic:true, color:{ argb: excelArgb('#677592') } };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
    cell.alignment = { vertical:'middle', horizontal:'center' };
    ws.getRow(6).height = 24;
    return;
  }
  addDetailExcelTable(ws, comiteRows, 'SometerComiteGerencia', 6);
}

function buildGerenciaEstadoWorkbook(rows){
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Forecast 2026 Provexpress';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.company = 'Provexpress';
  workbook.subject = 'Detalle negocios por estado';
  workbook.title = 'Forecast 2026 - Detalle negocios por estado';
  workbook.calcProperties = { fullCalcOnLoad: true };

  addResumenExcelSheet(workbook, rows);
  addComiteGerenciaExcelSheet(workbook, rows);
  addDetalleGeneralExcelSheet(workbook, rows);
  GERENCIA_ESTADOS.forEach(estado => {
    const stateRows = rows.filter(row => cleanDisplayText(row['ESTADO'], '').toUpperCase() === estado);
    addEstadoExcelSheet(workbook, estado, stateRows);
  });
  return workbook;
}

function loadExcelJsForExport(){
  if(typeof ExcelJS !== 'undefined') return Promise.resolve();
  if(EXCELJS_EXPORT_LOAD_PROMISE) return EXCELJS_EXPORT_LOAD_PROMISE;
  EXCELJS_EXPORT_LOAD_PROMISE = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://unpkg.com/exceljs@4.4.0/dist/exceljs.min.js';
    script.async = true;
    script.onload = () => {
      if(typeof ExcelJS !== 'undefined') resolve();
      else reject(new Error('ExcelJS no quedo disponible despues de cargar el respaldo.'));
    };
    script.onerror = () => {
      EXCELJS_EXPORT_LOAD_PROMISE = null;
      reject(new Error('No se pudo cargar ExcelJS.'));
    };
    document.head.appendChild(script);
  });
  return EXCELJS_EXPORT_LOAD_PROMISE;
}

function setGerenciaEstadoExportButtonBusy(isBusy){
  const btn = document.getElementById('btn-export-gerencia-estado');
  if(!btn) return;
  btn.disabled = Boolean(isBusy);
  btn.innerHTML = isBusy
    ? '<span class="export-excel-icon" aria-hidden="true">...</span><span>Generando</span>'
    : '<span class="export-excel-icon" aria-hidden="true">↓</span><span>Excel</span>';
}

async function downloadGerenciaEstadoExcel(){
  const rows = getGerenciaEstadoExportData();
  if(!rows.length){
    alert('No hay negocios para descargar con los filtros actuales.');
    return;
  }

  setGerenciaEstadoExportButtonBusy(true);
  try{
    await loadExcelJsForExport();
    const workbook = buildGerenciaEstadoWorkbook(rows);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Forecast_Detalle_Negocios_Estado_${getBogotaTimestampForFile()}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  } catch(error){
    console.error('[EXPORT GERENCIA ESTADO]', error);
    alert('No se pudo generar el Excel. Intenta nuevamente o revisa la consola para mas detalle.');
  } finally {
    setGerenciaEstadoExportButtonBusy(false);
  }
}

function getRowCostOriginalValue(row){
  return parseMonto(firstFilled(row, ['COSTO NEGOCIO','COSTO'])) || 0;
}

function getRowCostCopValue(row){
  const moneda = cleanDisplayText(row['MONEDA 2'], 'COP').trim().toUpperCase();
  const costo = getRowCostOriginalValue(row);
  return moneda === 'USD' ? costo * getTRM() : costo;
}

function getRowUtilidadCopValue(row){
  const utilidad = getUtilidad(row);
  return utilidad.moneda === 'USD' ? utilidad.valor * getTRM() : utilidad.valor;
}

function getMarcasLineasExportData(){
  const filters = getVisualCrossfilters('marcas');
  let rows = getVisibleMarcasData().slice();
  if(filters.marca) rows = rows.filter(row => getRowBrandName(row) === filters.marca);
  if(filters.linea) rows = rows.filter(row => getRowLineName(row) === filters.linea);

  const directorSelect = document.getElementById('sel-marca-director');
  const selectedDirector = directorSelect ? directorSelect.value || '' : '';
  if(selectedDirector) rows = rows.filter(row => cleanDisplayText(row['DIRECTOR'], '') === selectedDirector);

  return rows.sort((a,b) => toCOP(b) - toCOP(a));
}

function getMarcasLineasFilterText(){
  const filters = getVisualCrossfilters('marcas');
  const parts = [`Vista: ${CURRENT_USER ? getRoleLabel(CURRENT_USER.role) : 'Sin usuario'}`];
  if(isOscarMarcasGlobalScope()) parts.push('Alcance: Todos los usuarios');
  if(filters.marca) parts.push(`Marca: ${filters.marca}`);
  if(filters.linea) parts.push(`Linea: ${filters.linea}`);
  const directorSelect = document.getElementById('sel-marca-director');
  if(directorSelect && directorSelect.value) parts.push(`Director: ${directorSelect.value}`);
  return parts.join(' | ') || 'Sin filtros';
}

function getTopGroupedLabel(rows, getter, fallback){
  if(!(rows || []).length) return '';
  const totals = new Map();
  (rows || []).forEach(row => {
    const label = cleanDisplayText(getter(row), fallback || 'Sin dato');
    totals.set(label, (totals.get(label) || 0) + toCOP(row));
  });
  const top = [...totals.entries()].sort((a,b) => b[1] - a[1])[0];
  return top ? top[0] : (fallback || 'Sin dato');
}

function getMarcasTotalUSD(rows){
  return (rows || [])
    .filter(row => cleanDisplayText(row['MONEDA 2'], 'COP').trim().toUpperCase() === 'USD')
    .reduce((sum,row) => sum + (parseMonto(row['MONTO VENTA CLIENTE']) || 0), 0);
}

function getMarcasCategorySummaryRows(rows, type){
  const getter = type === 'marca'
    ? row => getRowBrandName(row)
    : row => getRowLineName(row);
  const otherGetter = type === 'marca'
    ? row => getRowLineName(row)
    : row => getRowBrandName(row);
  const fallback = type === 'marca' ? 'Sin marca' : 'Sin linea';
  const otherFallback = type === 'marca' ? 'Sin linea' : 'Sin marca';
  const grandTotal = rows.reduce((sum,row) => sum + toCOP(row), 0);
  const labels = [...new Set(rows.map(row => cleanDisplayText(getter(row), fallback)))].sort();

  return labels.map(label => {
    const groupRows = rows.filter(row => cleanDisplayText(getter(row), fallback) === label);
    const totalCOP = groupRows.reduce((sum,row) => sum + toCOP(row), 0);
    const costoCOP = groupRows.reduce((sum,row) => sum + getRowCostCopValue(row), 0);
    const utilidadCOP = groupRows.reduce((sum,row) => sum + getRowUtilidadCopValue(row), 0);
    return [
      label,
      groupRows.length,
      totalCOP,
      getMarcasTotalUSD(groupRows),
      costoCOP,
      utilidadCOP,
      grandTotal ? totalCOP / grandTotal : 0,
      getTopGroupedLabel(groupRows, otherGetter, otherFallback)
    ];
  }).sort((a,b) => b[2] - a[2]);
}

function getMarcasEstadoSummaryRows(rows){
  return GERENCIA_ESTADOS.map(estado => {
    const stateRows = rows.filter(row => cleanDisplayText(row['ESTADO'], '').toUpperCase() === estado);
    return [
      estado,
      stateRows.length,
      stateRows.reduce((sum,row) => sum + toCOP(row), 0),
      getMarcasTotalUSD(stateRows),
      stateRows.reduce((sum,row) => sum + getRowCostCopValue(row), 0),
      stateRows.reduce((sum,row) => sum + getRowUtilidadCopValue(row), 0),
      getTopGroupedLabel(stateRows, row => getRowBrandName(row), 'Sin marca'),
      getTopGroupedLabel(stateRows, row => getRowLineName(row), 'Sin linea')
    ];
  });
}

function getMarcasExecSummaryRows(rows){
  const execs = [...new Set(rows.map(row => cleanDisplayText(row['COMERCIAL'], 'Sin ejecutivo')))].sort();
  return execs.map(ejecutivo => {
    const execRows = rows.filter(row => cleanDisplayText(row['COMERCIAL'], 'Sin ejecutivo') === ejecutivo);
    const director = cleanDisplayText((execRows[0] && execRows[0]['DIRECTOR']) || '', 'Sin director');
    return [
      ejecutivo,
      director,
      getTopGroupedLabel(execRows, row => getRowBrandName(row), 'Sin marca'),
      getTopGroupedLabel(execRows, row => getRowLineName(row), 'Sin linea'),
      execRows.length,
      execRows.reduce((sum,row) => sum + toCOP(row), 0),
      getMarcasTotalUSD(execRows),
      execRows.reduce((sum,row) => sum + getRowCostCopValue(row), 0),
      execRows.reduce((sum,row) => sum + getRowUtilidadCopValue(row), 0)
    ];
  }).sort((a,b) => b[5] - a[5]);
}

function formatMarcasEstadoSummaryColumns(ws, startRow, rowCount){
  const totalRow = startRow + rowCount + 1;
  for(let rowNumber = startRow + 1; rowNumber <= totalRow; rowNumber++){
    ws.getCell(rowNumber, 2).numFmt = '#,##0';
    ws.getCell(rowNumber, 3).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 4).numFmt = excelMoneyFormat('USD', false);
    ws.getCell(rowNumber, 5).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 6).numFmt = excelMoneyFormat('COP', false);
    [2,3,4,5,6].forEach(col => {
      ws.getCell(rowNumber, col).alignment = { vertical:'top', horizontal:'right' };
    });
  }
}

function formatMarcasCategorySummaryColumns(ws, startRow, rowCount){
  const totalRow = startRow + rowCount + 1;
  for(let rowNumber = startRow + 1; rowNumber <= totalRow; rowNumber++){
    ws.getCell(rowNumber, 2).numFmt = '#,##0';
    ws.getCell(rowNumber, 3).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 4).numFmt = excelMoneyFormat('USD', false);
    ws.getCell(rowNumber, 5).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 6).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 7).numFmt = '0.##%';
    [2,3,4,5,6,7].forEach(col => {
      ws.getCell(rowNumber, col).alignment = { vertical:'top', horizontal:'right' };
    });
  }
}

function formatMarcasExecSummaryColumns(ws, startRow, rowCount){
  const totalRow = startRow + rowCount + 1;
  for(let rowNumber = startRow + 1; rowNumber <= totalRow; rowNumber++){
    ws.getCell(rowNumber, 5).numFmt = '#,##0';
    ws.getCell(rowNumber, 6).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 7).numFmt = excelMoneyFormat('USD', false);
    ws.getCell(rowNumber, 8).numFmt = excelMoneyFormat('COP', false);
    ws.getCell(rowNumber, 9).numFmt = excelMoneyFormat('COP', false);
    [5,6,7,8,9].forEach(col => {
      ws.getCell(rowNumber, col).alignment = { vertical:'top', horizontal:'right' };
    });
  }
}

function addMarcasLineasResumenSheet(workbook, rows){
  const ws = workbook.addWorksheet('Resumen marcas lineas');
  setupExcelWorksheet(ws, 8, 9);
  styleExcelTitle(
    ws,
    'Forecast 2026 - Marcas & Lineas',
    `Generado: ${getBogotaTimestampLabel()} | Registros exportados: ${rows.length}`,
    getMarcasLineasFilterText(),
    8
  );
  styleExcelKpis(ws, rows);
  styleExcelSectionLabel(ws, 8, 'Resumen por estado', 8);

  const columns = [
    { name:'Estado', totalsRowLabel:'Total' },
    { name:'Negocios', totalsRowFunction:'sum' },
    { name:'Total COP', totalsRowFunction:'sum' },
    { name:'Total USD original', totalsRowFunction:'sum' },
    { name:'Costo COP liq.', totalsRowFunction:'sum' },
    { name:'Utilidad COP liq.', totalsRowFunction:'sum' },
    { name:'Top Marca' },
    { name:'Top Linea' }
  ];
  const summaryRows = getMarcasEstadoSummaryRows(rows);
  [16,12,18,18,18,18,22,22].forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });
  addExcelTable(ws, {
    name: 'ResumenEstadosMarcas',
    ref: 'A9',
    columns,
    rows: summaryRows,
    theme: 'TableStyleMedium4'
  });
  styleExcelTable(ws, 9, summaryRows.length, columns.length, {
    statusColumn: 1,
    rightColumns: [2,3,4,5,6]
  });
  formatMarcasEstadoSummaryColumns(ws, 9, summaryRows.length);
}

function addMarcasCategoryExcelSheet(workbook, rows, type){
  const typeLabel = type === 'marca' ? 'Marca' : 'Linea';
  const otherLabel = type === 'marca' ? 'Top Linea' : 'Top Marca';
  const sheetName = type === 'marca' ? 'Resumen marcas' : 'Resumen lineas';
  const tableName = type === 'marca' ? 'ResumenPorMarca' : 'ResumenPorLinea';
  const ws = workbook.addWorksheet(sheetName);
  setupExcelWorksheet(ws, 8, 9);
  styleExcelTitle(
    ws,
    `Forecast 2026 - Resumen por ${typeLabel}`,
    `Generado: ${getBogotaTimestampLabel()} | Registros exportados: ${rows.length}`,
    getMarcasLineasFilterText(),
    8
  );
  styleExcelKpis(ws, rows);
  styleExcelSectionLabel(ws, 8, `Resumen por ${typeLabel.toLowerCase()}`, 8);

  const columns = [
    { name:typeLabel, totalsRowLabel:'Total' },
    { name:'Negocios', totalsRowFunction:'sum' },
    { name:'Total COP', totalsRowFunction:'sum' },
    { name:'Total USD original', totalsRowFunction:'sum' },
    { name:'Costo COP liq.', totalsRowFunction:'sum' },
    { name:'Utilidad COP liq.', totalsRowFunction:'sum' },
    { name:'Participacion', totalsRowFunction:'sum' },
    { name:otherLabel }
  ];
  const summaryRows = getMarcasCategorySummaryRows(rows, type);
  [24,12,18,18,18,18,14,24].forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });
  addExcelTable(ws, {
    name: tableName,
    ref: 'A9',
    columns,
    rows: summaryRows,
    theme: type === 'marca' ? 'TableStyleMedium9' : 'TableStyleMedium2'
  });
  styleExcelTable(ws, 9, summaryRows.length, columns.length, {
    rightColumns: [2,3,4,5,6,7]
  });
  formatMarcasCategorySummaryColumns(ws, 9, summaryRows.length);
}

function addMarcasEjecutivoExcelSheet(workbook, rows){
  const ws = workbook.addWorksheet('Marca por ejecutivo');
  setupExcelWorksheet(ws, 9, 9);
  styleExcelTitle(
    ws,
    'Forecast 2026 - Marca por ejecutivo',
    `Generado: ${getBogotaTimestampLabel()} | Registros exportados: ${rows.length}`,
    getMarcasLineasFilterText(),
    9
  );
  styleExcelKpis(ws, rows);
  styleExcelSectionLabel(ws, 8, 'Top marca y linea por ejecutivo', 9);

  const columns = [
    { name:'Ejecutivo', totalsRowLabel:'Total' },
    { name:'Director' },
    { name:'Top Marca' },
    { name:'Top Linea' },
    { name:'Negocios', totalsRowFunction:'sum' },
    { name:'Total COP', totalsRowFunction:'sum' },
    { name:'Total USD original', totalsRowFunction:'sum' },
    { name:'Costo COP liq.', totalsRowFunction:'sum' },
    { name:'Utilidad COP liq.', totalsRowFunction:'sum' }
  ];
  const execRows = getMarcasExecSummaryRows(rows);
  [26,24,22,22,12,18,18,18,18].forEach((width, idx) => { ws.getColumn(idx + 1).width = width; });
  addExcelTable(ws, {
    name: 'MarcaPorEjecutivo',
    ref: 'A9',
    columns,
    rows: execRows,
    theme: 'TableStyleMedium4'
  });
  styleExcelTable(ws, 9, execRows.length, columns.length, {
    rightColumns: [5,6,7,8,9]
  });
  formatMarcasExecSummaryColumns(ws, 9, execRows.length);
}

function addMarcasComiteGerenciaExcelSheet(workbook, rows){
  const comiteRows = getGerenciaComiteRows(rows);
  const ws = workbook.addWorksheet('Someter comite gerencia general');
  setupExcelWorksheet(ws, GERENCIA_EXCEL_DETAIL_COLS, 6);
  styleExcelTitle(
    ws,
    'Someter a comite de gerencia general',
    `Negocios mayores a ${fmtCOP(GERENCIA_COMITE_THRESHOLD_COP)} | Registros: ${comiteRows.length}`,
    getMarcasLineasFilterText(),
    GERENCIA_EXCEL_DETAIL_COLS
  );
  if(!comiteRows.length){
    ws.mergeCells(`A6:${excelColumnLetter(GERENCIA_EXCEL_DETAIL_COLS)}6`);
    const cell = ws.getCell('A6');
    cell.value = 'No hay negocios mayores a $1.000 millones COP con los filtros actuales.';
    cell.font = { name:'Aptos', size:11, italic:true, color:{ argb: excelArgb('#677592') } };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: excelArgb('#FAFCFF') } };
    cell.alignment = { vertical:'middle', horizontal:'center' };
    ws.getRow(6).height = 24;
    return;
  }
  addDetailExcelTable(ws, comiteRows, 'SometerComiteMarcas', 6);
}

function addMarcasDetalleExcelSheet(workbook, rows){
  const ws = workbook.addWorksheet('Detalle marcas lineas');
  setupExcelWorksheet(ws, GERENCIA_EXCEL_DETAIL_COLS, 6);
  styleExcelTitle(
    ws,
    'Detalle marcas y lineas',
    `Generado: ${getBogotaTimestampLabel()} | Registros: ${rows.length}`,
    getMarcasLineasFilterText(),
    GERENCIA_EXCEL_DETAIL_COLS
  );
  addDetailExcelTable(ws, rows, 'DetalleMarcasLineas', 6);
}

function buildMarcasLineasWorkbook(rows){
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Forecast 2026 Provexpress';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.company = 'Provexpress';
  workbook.subject = 'Marcas y lineas';
  workbook.title = 'Forecast 2026 - Marcas y lineas';
  workbook.calcProperties = { fullCalcOnLoad: true };

  addMarcasLineasResumenSheet(workbook, rows);
  addMarcasCategoryExcelSheet(workbook, rows, 'marca');
  addMarcasCategoryExcelSheet(workbook, rows, 'linea');
  addMarcasEjecutivoExcelSheet(workbook, rows);
  addMarcasComiteGerenciaExcelSheet(workbook, rows);
  addMarcasDetalleExcelSheet(workbook, rows);
  return workbook;
}

function setMarcasLineasExportButtonBusy(isBusy){
  const btn = document.getElementById('btn-export-marcas-lineas');
  if(!btn) return;
  btn.disabled = Boolean(isBusy);
  btn.innerHTML = isBusy
    ? '<span class="export-excel-icon" aria-hidden="true">...</span><span>Generando</span>'
    : '<span class="export-excel-icon" aria-hidden="true">↓</span><span>Excel</span>';
}

async function downloadMarcasLineasExcel(){
  const rows = getMarcasLineasExportData();
  if(!rows.length){
    alert('No hay negocios de marcas y lineas para descargar con los filtros actuales.');
    return;
  }

  setMarcasLineasExportButtonBusy(true);
  try{
    await loadExcelJsForExport();
    const workbook = buildMarcasLineasWorkbook(rows);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Forecast_Marcas_Lineas_${getBogotaTimestampForFile()}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  } catch(error){
    console.error('[EXPORT MARCAS LINEAS]', error);
    alert('No se pudo generar el Excel de marcas y lineas. Intenta nuevamente o revisa la consola para mas detalle.');
  } finally {
    setMarcasLineasExportButtonBusy(false);
  }
}

/* ══════════════════════════════════════
   DIRECTOR
══════════════════════════════════════ */
function renderDirector(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const dir=document.getElementById('sel-director').value;
  const directorMonthRows = dir ? ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir) : ALL_DATA;
  const mes=syncMonthSelectOptions('sel-dir-mes', getForecastMonths(directorMonthRows));
  const est=document.getElementById('sel-dir-estado').value;
  const trm=getTRM();
  const estadoOptions = [
    { value:'', label:'Todos' },
    { value:'GANADA', label:'GANADA' },
    { value:'PENDIENTE', label:'PENDIENTE' },
    { value:'PERDIDA', label:'PERDIDA' },
    { value:'APLAZADO', label:'APLAZADO' }
  ];
  const detailToolbar = `
    <div class="director-table-toolbar">
      <div>
        <div class="director-table-toolbar-label">Filtro rapido de tabla</div>
        <div class="director-table-toolbar-meta">Estado sincronizado con el filtro superior del director</div>
      </div>
      <div class="filter-group director-table-filter">
        <div class="filter-label">Estado tabla</div>
        <select id="sel-dir-estado-detail" onchange="setDirectorEstadoFilter(this.value)">
          ${buildOptionList(estadoOptions, { getValue: opt => opt.value, getLabel: opt => opt.label, selectedValue: est })}
        </select>
      </div>
    </div>`;

  let data=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir);
  if(mes) data=data.filter(r=>getMonth(getRowDateValue(r))===mes);
  if(est) data=data.filter(r=>r['ESTADO']===est);
  
  const totalCOP=data.reduce((s,r)=>s+toCOP(r),0);
  const totalUSD=data.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const utilidadCOP = sumUtilidad(data, 'COP');
  const utilidadUSD = sumUtilidad(data, 'USD');
  const ganadas=data.filter(r=>r['ESTADO']==='GANADA');
  // Todos los ejecutivos de este director (con o sin datos)
  const execsWithDataDir=[...new Set(data.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  const execsFromFilesDir=(LOADED_FILES_BY_DIR[dir]||[]).map(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim()).filter(Boolean);
  const execs=[...new Set([...execsWithDataDir,...execsFromFilesDir])].sort();
  
  // Evolution del director
  const evoMonthKeys = getForecastMonths(data.length ? data : ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir));
  const evoMonths = {};
  evoMonthKeys.forEach(m => { evoMonths[m] = 0; });
  data.forEach(r=>{const m=getMonth(getRowDateValue(r));if(evoMonths[m]!==undefined)evoMonths[m]+=toCOP(r);});
  
  const evoSVG=()=>{
    const months=Object.keys(evoMonths);
    if(!months.length) return `<div style="font-size:11px;color:var(--text2)">Sin datos mensuales para mostrar.</div>`;
    const vals=Object.values(evoMonths);
    const maxV=Math.max(...vals,1);
    const W=400,H=140,padL=44,padR=10,padT=16,padB=28;
    const gW=W-padL-padR,gH=H-padT-padB;
    const bW=Math.min(36,(gW/months.length)-12);
    let s=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;overflow:visible">`;
    s+=`<defs><linearGradient id="bg-dir" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2D7BF7" stop-opacity=".9"/><stop offset="100%" stop-color="#2D7BF7" stop-opacity=".35"/></linearGradient></defs>`;
    // Grid
    [0,.5,1].forEach(t=>{
      const y=padT+gH*(1-t);
      s+=`<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="var(--border)" stroke-width="${t===0?1.5:.6}"/>`;
      if(t>0) s+=`<text x="${padL-4}" y="${y+3}" text-anchor="end" font-size="7.5" fill="var(--text3)" font-weight="400" font-family="IBM Plex Mono,monospace">${abr(maxV*t)}</text>`;
    });
    months.forEach((m,i)=>{
      const v=vals[i]||0;
      const bh=Math.max(v/maxV*gH, v>0?2:0);
      const cx=padL+(i+0.5)*(gW/months.length);
      const x=cx-bW/2;
      const y=padT+gH-bh;
      const monthClick = jsCall('toggleDirectorMonthFilter', m);
      s+=`<rect class="evo-month-bar ${mes===m?'chart-selected':''}" x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bW}" height="${bh.toFixed(1)}" rx="3" fill="url(#bg-dir)" role="button" tabindex="0" onclick="${escAttr(monthClick)}" onkeydown="${escAttr(`if(event.key==='Enter'||event.key===' '){event.preventDefault();${monthClick}}`)}" data-tooltip="${escAttr('Filtrar mes '+getMonthShortLabel(m)+': '+abr(v))}"></rect>`;
      s+=`<text x="${cx.toFixed(1)}" y="${H-6}" text-anchor="middle" font-size="9" fill="var(--text3)" font-family="IBM Plex Sans,sans-serif" font-weight="400">${getMonthShortLabel(m)}</text>`;
    });
    s+='</svg>';
    return s;
  };
  
  // Equipo execs
  const ejCards=execs.map((e,i)=>{
    const ejData=data.filter(r=>(r['COMERCIAL']||'').trim()===e);
    const ejCOP=ejData.reduce((s,r)=>s+toCOP(r),0);
    const ejGan=ejData.filter(r=>r['ESTADO']==='GANADA').length;
    const hasD=ejData.length>0;
    const c=COLORS[i%COLORS.length];
    const ini=e.split(' ').slice(0,2).map(w=>w[0]).join('');
    const isSelected=(SELECTED_EXEC_BY_DIR[dir]||'')===e;
    return `<div class="kpi exec-card ${isSelected?'selected':''}" style="--ac:${c};min-width:160px;opacity:${hasD?1:.55}" onclick="${escAttr(jsCall('selectDirectorExec', e))}">
      <div class="kpi-accent"></div>
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
        <div style="width:30px;height:30px;border-radius:50%;background:${c}30;border:1px solid ${c}60;display:flex;align-items:center;justify-content:center;font-family:var(--font-display);font-size:11px;font-weight:700;color:${c}">${escHtml(ini)}</div>
        <div>
          <div style="font-size:11px;font-family:var(--font-display);font-weight:700;color:var(--text)">${escHtml(e.split(' ')[0])}</div>
          <div style="font-size:9px;color:var(--text3)">${hasD?ejData.length+' negocios':'Sin datos aún'}</div>
        </div>
      </div>
      ${renderLastConnection(e)}
      <div class="kpi-val">${hasD?abr(ejCOP):'—'}</div>
      <div class="kpi-sub">${hasD?ejGan+' ganadas':'Pendiente'}</div>
    </div>`;
  }).join('');

  const selectedExec = SELECTED_EXEC_BY_DIR[dir] || '';
  const execData = selectedExec ? data.filter(r=>(r['COMERCIAL']||'').trim()===selectedExec) : [];
  const execTotalCOP = execData.reduce((s,r)=>s+toCOP(r),0);
  const execTotalUSD = execData.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const execUtilidadCOP = sumUtilidad(execData, 'COP');
  const execUtilidadUSD = sumUtilidad(execData, 'USD');
  const execGanadas = execData.filter(r=>r['ESTADO']==='GANADA');
  const execKPIs = selectedExec ? `
    <div class="section-hd" style="margin-top:14px">
      <h2>${escHtml(selectedExec)}</h2>
      <span class="section-tag">EJECUTIVO</span>
      <button class="btn-clear" onclick="clearDirectorExec()">Ver todo el equipo</button>
    </div>
    <div class="kpi-grid kpi-grid-6" style="margin-bottom:16px">
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(execTotalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(execTotalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        ${kpiUSD(execTotalUSD)}
        <div class="kpi-sub">Liq: ${abr(execTotalUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad COP</div>
        <div class="kpi-val">${abr(execUtilidadCOP)}</div>
        <div class="kpi-sub">${fmtCOP(execUtilidadCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad USD</div>
        ${kpiUSD(execUtilidadUSD)}
        <div class="kpi-sub">Liq: ${abr(execUtilidadUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ganadas</div>
        <div class="kpi-val">${execGanadas.length}</div>
        <div class="kpi-sub">${abr(execGanadas.reduce((s,r)=>s+toCOP(r),0))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Negocios</div>
        <div class="kpi-val">${execData.length}</div>
        <div class="kpi-sub">Director: ${escHtml(dir)}</div>
      </div>
    </div>
    <div class="chart-card g1">
      <div class="chart-hd">Detalle de Negocios — ${escHtml(selectedExec)}</div>
      ${detailToolbar}
      ${execData.length ? buildTable(execData, { clickable:true, sourcePage:'director' }) : `<div style="font-size:11px;color:var(--text2)">Sin registros para este filtro.</div>`}
    </div>
  ` : `
    <div class="chart-card g1">
      <div class="chart-hd">Detalle de Negocios</div>
      ${detailToolbar}
      ${buildTable(data.slice(0,30), { clickable:true, sourcePage:'director' })}
    </div>
  `;
  
  document.getElementById('director-content').innerHTML=`
    <div class="section-hd"><h2>${escHtml(dir)}</h2><span class="section-tag">DIRECTOR</span></div>
    
    <div class="kpi-grid kpi-grid-6" style="margin-bottom:16px">
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        ${kpiUSD(totalUSD)}
        <div class="kpi-sub">Liq: ${abr(totalUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad COP</div>
        <div class="kpi-val">${abr(utilidadCOP)}</div>
        <div class="kpi-sub">${fmtCOP(utilidadCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad USD</div>
        ${kpiUSD(utilidadUSD)}
        <div class="kpi-sub">Liq: ${abr(utilidadUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ganadas</div>
        <div class="kpi-val">${ganadas.length}</div>
        <div class="kpi-sub">${abr(ganadas.reduce((s,r)=>s+toCOP(r),0))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ejecutivos</div>
        <div class="kpi-val">${execs.length}</div>
        <div class="kpi-sub">${data.length} negocios · ${execsWithDataDir.length} activos</div>
      </div>
    </div>

    <div class="g2b">
      <div class="chart-card">
        <div class="chart-hd">Evolución Mensual — ${escHtml(dir.split(' ')[0])}</div>
        ${evoSVG()}
      </div>
      <div class="chart-card">
        <div class="chart-hd">Estado Pipeline</div>
        <div class="donut-wrap">
          <svg id="donut-dir-est" viewBox="0 0 100 100" style="width:130px;height:130px;flex-shrink:0"></svg>
          <div class="donut-leg" id="leg-dir-est"></div>
        </div>
      </div>
    </div>

    <div class="section-hd"><h2>Equipo</h2><span class="section-tag">${execs.length} EJECUTIVOS</span>${buildExcelConnectionExportButton('director', 'Excel')}</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:18px">
      ${ejCards}
    </div>
    ${execKPIs}
  `;
  attachChartTooltips(document.getElementById('director-content'));
  
  // Donut estado director — usar datos sin filtro de estado para mostrar distribución real
  const dataForDonut=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir && (!mes||getMonth(getRowDateValue(r))===mes));
  const estD=['GANADA','PENDIENTE','PERDIDA','APLAZADO'].map(e=>({name:e,val:dataForDonut.filter(r=>r['ESTADO']===e).reduce((s,r)=>s+toCOP(r),0)}));
  renderDonut('donut-dir-est','leg-dir-est',estD,{
    hasActive: Boolean(est),
    getOnClick: item => jsCall('toggleDirectorEstadoFilter', item.name),
    getIsSelected: item => est === item.name,
    clickTitle: 'Filtrar por estado',
    tooltipPrefix: 'Filtrar estado: '
  });
}

function selectDirectorExec(name){
  const dir=document.getElementById('sel-director').value;
  if(!dir) return;
  SELECTED_EXEC_BY_DIR[dir]=name;
  renderDirector();
}

function clearDirectorExec(){
  const dir=document.getElementById('sel-director').value;
  if(dir) delete SELECTED_EXEC_BY_DIR[dir];
  renderDirector();
}

function setDirectorEstadoFilter(value){
  const topFilter = document.getElementById('sel-dir-estado');
  if(topFilter) topFilter.value = value;
  renderDirector();
}

/* ══════════════════════════════════════
   EJECUTIVO
══════════════════════════════════════ */
function initials(name){return name.split(' ').slice(0,2).map(w=>w[0]).join('');}

function renderEjecutivo(){
  const ALL_DATA = EJECUTIVO_BRAND_FOCUS ? getVisibleMarcasData() : getVisibleData();
  if(!ALL_DATA.length) return;
  const role = CURRENT_USER ? CURRENT_USER.role : null;
  const targetName = role === 'ejecutivo' ? getExecTargetName() : '';
  const selEj = document.getElementById('sel-ejecutivo');
  const est=document.getElementById('sel-ej-estado').value;
  const trm=getTRM();
  
  // Persona grid — todos los ejecutivos cargados, tengan o no datos
  const execsFromData = [...new Set(ALL_DATA.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  // Agregar ejecutivos de archivos cargados aunque estén vacíos
  const execsFromFiles = Object.values(LOADED_FILES_BY_DIR||{}).flat()
    .map(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim());
  const allExecs = [...new Set([...execsFromData, ...execsFromFiles])].sort();
  if(selEj){
    if(role === 'ejecutivo' && targetName){
      selEj.innerHTML = optionHtml(targetName, targetName, false);
      selEj.value = targetName;
    } else {
      const current = selEj.value;
      selEj.innerHTML = buildOptionList(allExecs);
      if(current && allExecs.includes(current)) selEj.value = current;
      else if(allExecs[0]) selEj.value = allExecs[0];
    }
  }
  const ej = (role === 'ejecutivo' && targetName) ? targetName : (selEj ? selEj.value : '');
  const execs = (role === 'ejecutivo' && targetName) ? [targetName] : allExecs;
  if(EJECUTIVO_BRAND_FOCUS && !namesMatch(EJECUTIVO_BRAND_FOCUS.execName, ej)) {
    EJECUTIVO_BRAND_FOCUS = null;
    renderEjecutivo();
    return;
  }
  const focusedBrand = EJECUTIVO_BRAND_FOCUS ? EJECUTIVO_BRAND_FOCUS.brandName : '';
  const focusedDirector = EJECUTIVO_BRAND_FOCUS ? EJECUTIVO_BRAND_FOCUS.directorName : '';
  const brandKey = normalizeCategoryValue(focusedBrand);
  let execMonthRows = ej ? ALL_DATA.filter(r=>namesMatch(r['COMERCIAL'], ej)) : ALL_DATA;
  if(focusedDirector) execMonthRows = execMonthRows.filter(r => cleanDisplayText(r['DIRECTOR'], '') === focusedDirector);
  if(focusedBrand) execMonthRows = execMonthRows.filter(r => normalizeCategoryValue(getRowBrandName(r)) === brandKey);
  const mes=syncMonthSelectOptions('sel-ej-mes', getForecastMonths(execMonthRows));
  document.getElementById('persona-grid').innerHTML=execs.map((e,i)=>{
    const ed=ALL_DATA.filter(r=>(r['COMERCIAL']||'').trim()===e);
    const cop=ed.reduce((s,r)=>s+toCOP(r),0);
    const gan=ed.filter(r=>r['ESTADO']==='GANADA').length;
    const pen=ed.filter(r=>r['ESTADO']==='PENDIENTE').length;
    const c=COLORS[i%COLORS.length];
    const dirFromData=ed[0]?ed[0]['DIRECTOR']||'':'';
    const dirFromFile=Object.entries(LOADED_FILES_BY_DIR||{}).find(([d,fs])=>fs.some(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim()===e));
    const dir=dirFromData||(dirFromFile?dirFromFile[0]:'—');
    const hasData=ed.length>0;
    const selected=e===ej?'selected':'';
    return `<div class="persona-card ${selected} ${hasData?'':'no-data'}" onclick="${escAttr(jsCall('selectEjecutivo', e))}">
      <div class="persona-avatar" style="background:${c}${hasData?'25':'10'};border:2px solid ${c}${hasData?'50':'20'};color:${hasData?c:'var(--text2)'}">${escHtml(initials(e))}</div>
      <div class="persona-name" style="color:${hasData?'var(--text)':'var(--text2)'}">${escHtml(e)}</div>
      <div class="persona-role">${escHtml(dir)}</div>
      ${renderLastConnection(e)}
      ${hasData
        ?`<div class="persona-stats">
        <div class="p-stat"><div class="p-stat-label">Total</div><div class="p-stat-val" style="color:${c};font-size:11px">${abr(cop)}</div></div>
        <div class="p-stat"><div class="p-stat-label">Negoc.</div><div class="p-stat-val">${ed.length}</div></div>
        <div class="p-stat"><div class="p-stat-label">Ganadas</div><div class="p-stat-val" style="color:var(--corp-green)">${gan}</div></div>
        <div class="p-stat"><div class="p-stat-label">Pend.</div><div class="p-stat-val" style="color:var(--corp-amber)">${pen}</div></div>
      </div>`
        :`<div style="font-size:9px;color:var(--text3);font-family:var(--font-display);margin-top:8px;padding:5px 8px;background:rgba(255,255,255,.03);border-radius:6px;letter-spacing:.5px">📋 Sin registros aún</div>`}
    </div>`;
  }).join('');
  
  if(!ej) return;
  
  let data=ALL_DATA.filter(r=>r['COMERCIAL']===ej);
  if(mes) data=data.filter(r=>getMonth(getRowDateValue(r))===mes);
  if(est) data=data.filter(r=>r['ESTADO']===est);
  if(focusedDirector) {
    data = data.filter(r => cleanDisplayText(r['DIRECTOR'], '') === focusedDirector);
  }
  if(focusedBrand) {
    data = data.filter(r => normalizeCategoryValue(getRowBrandName(r)) === brandKey);
  }
  const ejecutivoBaseData = data.slice();
  const ejecutivoFilters = getVisualCrossfilters('ejecutivo');
  if(ejecutivoFilters.linea) data = data.filter(r => getRowLineName(r) === ejecutivoFilters.linea);
  
  const totalCOP=data.reduce((s,r)=>s+toCOP(r),0);
  const totalUSD=data.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const ganadas=data.filter(r=>r['ESTADO']==='GANADA');
  const ejColor=COLORS[execs.indexOf(ej)%COLORS.length];
  const focusBadge = focusedBrand
    ? `<span class="section-tag" style="background:rgba(42,191,223,.16);color:var(--corp-cyan);border-color:rgba(42,191,223,.32)">MARCA · ${escHtml(focusedBrand)}</span>`
    : '';
  const directorFocusBadge = focusedDirector
    ? `<span class="section-tag">DIRECTOR · ${escHtml(focusedDirector)}</span>`
    : '';
  const focusClear = focusedBrand
    ? `<button class="btn-clear" onclick="clearEjecutivoBrandFocus()">Ver todo el ejecutivo</button>`
    : '';
  
  const linData=buildLineValueData(data);
  
  document.getElementById('ejecutivo-content').innerHTML=`
    <div class="section-hd" style="margin-top:16px"><h2>${escHtml(ej)}</h2><span class="section-tag" style="background:${ejColor}20;color:${ejColor};border-color:${ejColor}40">EJECUTIVO</span>${focusBadge}${directorFocusBadge}${focusClear}</div>
    ${buildVisualCrossfilterBar('ejecutivo', ejecutivoBaseData.length, data.length, [['linea','Linea']])}
    
    <div class="kpi-grid kpi-grid-4" style="margin-bottom:16px">
      <div class="kpi" style="--ac:${ejColor}"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        ${kpiUSD(totalUSD)}
        <div class="kpi-sub">Liq: ${abr(totalUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ganadas</div>
        <div class="kpi-val">${ganadas.length}</div>
        <div class="kpi-sub">${abr(ganadas.reduce((s,r)=>s+toCOP(r),0))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Negocios</div>
        <div class="kpi-val">${data.length}</div>
        <div class="kpi-sub">Director: ${escHtml((data[0]||{})['DIRECTOR']||'—')}</div>
      </div>
    </div>

    <div class="g2">
      <div class="chart-card">
        <div class="chart-hd">Top Líneas — ${escHtml(ej.split(' ')[0])}</div>
        <div class="bar-list" id="bar-ej-lineas"></div>
      </div>
      <div class="chart-card">
        <div class="chart-hd">Estado Negocios</div>
        <div class="donut-wrap">
          <svg id="donut-ej-est" viewBox="0 0 100 100" style="width:130px;height:130px;flex-shrink:0"></svg>
          <div class="donut-leg" id="leg-ej-est"></div>
        </div>
      </div>
    </div>
    
    <div class="chart-card g1">
      <div class="chart-hd">Detalle de Negocios — ${escHtml(ej)}${focusedBrand ? ` <span>${escHtml(focusedBrand)}</span>` : ''}</div>
      ${buildTable(data, { clickable:true, sourcePage:'ejecutivo' })}
    </div>
  `;
  
  renderBars('bar-ej-lineas',linData,COLORS,null,{
    getOnClick: item => jsCall('setVisualCrossfilter', 'ejecutivo', 'linea', item.name),
    getIsSelected: item => getVisualCrossfilters('ejecutivo').linea === item.name,
    clickTitle: 'Filtrar por linea',
    tooltipPrefix: 'Filtrar linea: '
  });
  const estD=['GANADA','PENDIENTE','PERDIDA','APLAZADO'].map(e=>({name:e,val:data.filter(r=>r['ESTADO']===e).length}));
  renderDonut('donut-ej-est','leg-ej-est',estD,{
    hasActive: Boolean(est),
    getOnClick: item => jsCall('toggleEjecutivoEstadoFilter', item.name),
    getIsSelected: item => est === item.name,
    clickTitle: 'Filtrar por estado',
    tooltipPrefix: 'Filtrar estado: '
  });
}

function selectEjecutivo(name){
  document.getElementById('sel-ejecutivo').value=name;
  if(EJECUTIVO_BRAND_FOCUS && !namesMatch(EJECUTIVO_BRAND_FOCUS.execName, name)) {
    EJECUTIVO_BRAND_FOCUS = null;
  }
  renderEjecutivo();
}

function clearEjecutivoBrandFocus(){
  EJECUTIVO_BRAND_FOCUS = null;
  renderEjecutivo();
}

function openExecNegociosFromMarcas(execName, directorName, brandName){
  const targetExec = cleanDisplayText(execName, '').trim();
  const targetBrand = cleanDisplayText(brandName, '').trim();
  if(!targetExec) return;
  const ejSelect = document.getElementById('sel-ejecutivo');
  const ejMonth = document.getElementById('sel-ej-mes');
  const ejEstado = document.getElementById('sel-ej-estado');
  if(ejSelect) ejSelect.value = targetExec;
  if(ejMonth) ejMonth.value = '';
  if(ejEstado) ejEstado.value = '';
  EJECUTIVO_BRAND_FOCUS = targetBrand ? {
    execName: targetExec,
    directorName: cleanDisplayText(directorName, '').trim(),
    brandName: targetBrand
  } : null;
  renderEjecutivo();
  showPage('ejecutivo', getNavButtonForPage('ejecutivo'));
  scrollToElementAfterRender('ejecutivo-content', 118);
}

function getMonthLabel(monthKey){
  if(!monthKey) return 'Sin mes';
  return getMonthLongLabel(monthKey);
}

function buildSalesTable(data, opts){
  const options = opts || {};
  const sourcePage = options.sourcePage || getActivePageId();
  return `<table class="responsive-table">
    <thead><tr><th>Fecha</th><th>Cotizacion</th><th>Cliente</th><th>Soporta</th><th>Moneda</th><th>Costo Negocio</th><th>Valor venta</th><th>Utilidad</th><th>Margen</th><th>TRM ref</th><th>Estado</th></tr></thead>
    <tbody>${data.length ? data.map(r=>{
      const mon = cleanDisplayText(r['MONEDA 2'], 'COP').trim().toUpperCase();
      const fecha = cleanDisplayText(formatDateValue(getRowDateValue(r)), 'Sin fecha');
      const cotizacion = cleanDisplayText(getSalesQuoteNumber(r), 'Sin numero');
      const cliente = cleanDisplayText(getRowClientName(r), 'Sin cliente');
      const soporta = cleanDisplayText(getSalesSoportaName(r), 'Sin apoyo');
      const costo = getSalesCostValue(r);
      const valor = parseMonto(r['MONTO VENTA CLIENTE']) || 0;
      const utilidad = getSalesUtilidadValue(r);
      const margen = formatMarginDisplay(r['MARGEN'], '-');
      const trmRef = parseMonto(r['TRM REFERENCIA']) || 0;
      const estado = cleanDisplayText(r['ESTADO'], 'Sin estado').toUpperCase();
      const estadoClass = getEstadoBadgeClass(estado);
      const rowAttrs = options.clickable && r.__RID
        ? ` class="table-row-action" onclick="${escAttr(jsCall('openNegocioDetailById', r.__RID, sourcePage))}" title="Abrir detalle del registro sales"`
        : '';
      return `<tr${rowAttrs}>
        <td class="td-mono" data-label="Fecha">${escHtml(fecha)}</td>
        <td class="td-mono" data-label="Cotizacion">${escHtml(cotizacion)}</td>
        <td data-label="Cliente">${escHtml(cliente)}</td>
        <td data-label="Soporta">${escHtml(soporta)}</td>
        <td data-label="Moneda"><span class="badge ${mon==='USD'?'badge-PEDIDA':'badge-PENDIENTE'}">${escHtml(mon)}</span></td>
        <td class="td-mono ${mon==='USD'?'td-usd':'td-cop'}" data-label="Costo Negocio">${mon==='USD'?fmtUSD(costo):fmtCOP(costo)}</td>
        <td class="td-mono ${mon==='USD'?'td-usd':'td-cop'}" data-label="Valor venta">${mon==='USD'?fmtUSD(valor):fmtCOP(valor)}</td>
        <td class="td-mono" style="color:var(--corp-cyan)" data-label="Utilidad">${mon==='USD'?fmtUSD(utilidad):fmtCOP(utilidad)}</td>
        <td class="td-mono" style="color:var(--corp-amber)" data-label="Margen">${escHtml(margen)}</td>
        <td class="td-mono" data-label="TRM ref">${trmRef ? fmtTRM(trmRef) : '—'}</td>
        <td data-label="Estado"><span class="badge badge-${estadoClass}">${escHtml(estado)}</span></td>
      </tr>`;
    }).join('') : `<tr><td colspan="11" style="text-align:center;color:var(--text2)">Sin registros sales para este filtro.</td></tr>`}</tbody>
  </table>`;
}

function getSalesViewMode(){
  return SALES_VIEW_MODE === 'pendientes' ? 'pendientes' : 'reporte';
}

function syncSalesViewButtons(){
  const mode = getSalesViewMode();
  const btnReporte = document.getElementById('btn-sales-reporte');
  const btnPendientes = document.getElementById('btn-sales-pendientes');
  if(btnReporte) btnReporte.classList.toggle('active', mode === 'reporte');
  if(btnPendientes) btnPendientes.classList.toggle('active', mode === 'pendientes');
}

function setSalesFilterVisibility(mode){
  const showReporte = mode === 'reporte';
  [
    ['sales-month-filter', showReporte],
    ['sales-state-filter', showReporte],
    ['sales-pending-type-filter', !showReporte],
    ['sales-pending-category-filter', !showReporte]
  ].forEach(([id, visible]) => {
    const el = document.getElementById(id);
    if(el) el.style.display = visible ? '' : 'none';
  });
}

function setSalesView(mode){
  SALES_VIEW_MODE = mode === 'pendientes' ? 'pendientes' : 'reporte';
  renderSales();
}

function setSalesPendingFilter(kind, value){
  const selPendiente = document.getElementById('sel-sales-pendiente');
  const selCategoria = document.getElementById('sel-sales-categoria');
  const type = kind || 'all';
  if(type === 'all') {
    if(selPendiente) selPendiente.value = '';
    if(selCategoria) selCategoria.value = '';
    SALES_PENDING_VALUE_ONLY = false;
  } else if(type === 'pendiente') {
    if(selPendiente) selPendiente.value = value || '';
    SALES_PENDING_VALUE_ONLY = false;
  } else if(type === 'categoria') {
    if(selCategoria) selCategoria.value = value || '';
    SALES_PENDING_VALUE_ONLY = false;
  } else if(type === 'valor') {
    SALES_PENDING_VALUE_ONLY = true;
  }
  SALES_VIEW_MODE = 'pendientes';
  renderSales();
}

function getSalesPendingBadgeClass(value){
  const key = normalizeCategoryValue(value);
  if(key.includes('pedido') || key.includes('compra')) return 'PEDIDA';
  if(key.includes('remision') || key.includes('factura')) return 'GANADA';
  return 'PENDIENTE';
}

function buildSalesPendingTable(data){
  return `<table class="responsive-table">
    <thead><tr><th>Pendiente</th><th>Numero pendiente</th><th>Cliente</th><th>Comercial</th><th>Categoria</th><th>Valor facturas</th></tr></thead>
    <tbody>${data.length ? data.map(r=>{
      const pendiente = getSalesPendingType(r);
      const numero = cleanDisplayText(getSalesPendingNumber(r), 'Sin numero');
      const cliente = cleanDisplayText(getRowClientName(r), 'Sin cliente');
      const comercial = cleanDisplayText(getSalesPendingCommercial(r), 'Sin comercial');
      const categoria = getSalesPendingCategory(r);
      const valor = getSalesPendingInvoiceValue(r);
      return `<tr>
        <td data-label="Pendiente"><span class="badge badge-${getSalesPendingBadgeClass(pendiente)}">${escHtml(pendiente)}</span></td>
        <td class="td-mono" data-label="Numero pendiente">${escHtml(numero)}</td>
        <td data-label="Cliente">${escHtml(cliente)}</td>
        <td data-label="Comercial">${escHtml(comercial)}</td>
        <td data-label="Categoria"><span class="badge badge-${getSalesPendingBadgeClass(categoria)}">${escHtml(categoria)}</span></td>
        <td class="td-mono td-cop" data-label="Valor facturas">${valor ? fmtCOP(valor) : '—'}</td>
      </tr>`;
    }).join('') : `<tr><td colspan="6" style="text-align:center;color:var(--text2)">Sin pendientes para este filtro.</td></tr>`}</tbody>
  </table>`;
}

function renderSales(){
  const allSales = getVisibleSalesData();
  const allPending = getVisibleSalesPendingData();
  const mode = getSalesViewMode();
  const role = CURRENT_USER ? CURRENT_USER.role : null;
  const targetName = role === 'sales_support' ? getSalesSupportTargetName() : '';
  const host = document.getElementById('sales-content');
  const grid = document.getElementById('sales-support-grid');
  const selSupport = document.getElementById('sel-sales-support');
  const selMes = document.getElementById('sel-sales-mes');
  const selEstado = document.getElementById('sel-sales-estado');
  const selPendiente = document.getElementById('sel-sales-pendiente');
  const selCategoria = document.getElementById('sel-sales-categoria');
  if(!host || !grid || !selSupport || !selMes || !selEstado) return;

  syncSalesViewButtons();
  setSalesFilterVisibility(mode);

  const supportsFromData = mode === 'pendientes'
    ? [...new Set(allPending.map(r=>getSalesPendingSupportName(r)).filter(Boolean))]
    : [...new Set(allSales.map(r=>getSalesSupportName(r)).filter(Boolean))];
  const supportsFromFiles = Object.keys(LOADED_SALES_BY_SUPPORT || {}).map(canonicalizeSalesSupportName).filter(Boolean);
  const allSupports = [...new Set([...supportsFromData, ...supportsFromFiles])].sort((a,b)=>a.localeCompare(b,'es'));

  if(role === 'sales_support' && targetName) {
    selSupport.innerHTML = optionHtml(targetName, targetName, false);
    selSupport.value = targetName;
  } else {
    const current = selSupport.value;
    selSupport.innerHTML = buildOptionList(allSupports);
    if(current && allSupports.includes(current)) selSupport.value = current;
    else if(allSupports[0]) selSupport.value = allSupports[0];
  }

  if(!allSupports.length){
    grid.innerHTML = '';
    host.innerHTML = `<div class="chart-card g1"><div style="font-size:12px;color:var(--text2)">No hay archivos Sales Support cargados para este usuario o grupo.</div></div>`;
    return;
  }

  const selectedSupport = role === 'sales_support' && targetName ? targetName : selSupport.value;

  if(mode === 'reporte') {
    const monthValues = [...new Set(allSales.map(r=>getMonth(getRowDateValue(r))).filter(Boolean))].sort();
    const currentMonth = selMes.value;
    selMes.innerHTML = optionHtml('', 'Todos', false) + buildOptionList(monthValues, {
      getLabel: getMonthLabel
    });
    if(currentMonth && monthValues.includes(currentMonth)) selMes.value = currentMonth;
    else selMes.value = '';
  } else if(selPendiente && selCategoria) {
    const supportPendingRows = allPending.filter(r => namesMatch(getSalesPendingSupportName(r), selectedSupport));
    const pendingTypes = [...new Set(supportPendingRows.map(getSalesPendingType).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
    const pendingCategories = [...new Set(supportPendingRows.map(getSalesPendingCategory).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
    const currentType = selPendiente.value;
    const currentCategory = selCategoria.value;
    selPendiente.innerHTML = optionHtml('', 'Todos', false) + buildOptionList(pendingTypes);
    selCategoria.innerHTML = optionHtml('', 'Todas', false) + buildOptionList(pendingCategories);
    if(currentType && pendingTypes.includes(currentType)) selPendiente.value = currentType;
    else selPendiente.value = '';
    if(currentCategory && pendingCategories.includes(currentCategory)) selCategoria.value = currentCategory;
    else selCategoria.value = '';
  }

  grid.innerHTML = allSupports.map((supportName, idx)=>{
    const supportRows = mode === 'pendientes'
      ? allPending.filter(r => namesMatch(getSalesPendingSupportName(r), supportName))
      : allSales.filter(r => namesMatch(getSalesSupportName(r), supportName));
    const totalCOP = mode === 'pendientes'
      ? supportRows.reduce((sum,row)=>sum+getSalesPendingInvoiceValue(row),0)
      : supportRows.reduce((sum,row)=>sum+toCOP(row),0);
    const totalGanadas = supportRows.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()==='GANADA').length;
    const supportedCount = mode === 'pendientes'
      ? [...new Set(supportRows.map(getSalesPendingCommercial).filter(Boolean))].length
      : [...new Set(supportRows.map(r=>getSalesSoportaName(r)).filter(Boolean))].length;
    const supportPedidoCount = mode === 'pendientes'
      ? supportRows.filter(row => normalizeCategoryValue(getSalesPendingType(row)).includes('pedido')).length
      : 0;
    const supportGlpiCount = mode === 'pendientes'
      ? supportRows.filter(row => normalizeCategoryValue(getSalesPendingType(row)).includes('glpi')).length
      : 0;
    const supportFacturaCount = mode === 'pendientes'
      ? supportRows.filter(row =>
          normalizeCategoryValue(getSalesPendingCategory(row)).includes('factura')
          || normalizeCategoryValue(getSalesPendingType(row)).includes('remision')
        ).length
      : 0;
    const c = COLORS[idx % COLORS.length];
    const hasData = supportRows.length > 0;
    const selected = namesMatch(selectedSupport, supportName) ? 'selected' : '';
    return `<div class="persona-card ${selected} ${hasData?'':'no-data'}" onclick="${escAttr(jsCall('selectSalesSupport', supportName))}">
      <div class="persona-avatar" style="background:${c}${hasData?'25':'10'};border:2px solid ${c}${hasData?'50':'20'};color:${hasData?c:'var(--text2)'}">${escHtml(initials(supportName))}</div>
      <div class="persona-name" style="color:${hasData?'var(--text)':'var(--text2)'}">${escHtml(supportName)}</div>
      <div class="persona-role">Sales Support</div>
      ${hasData
        ? (mode === 'pendientes'
          ? `<div class="persona-stats persona-stats-3">
              <div class="p-stat"><div class="p-stat-label">Facturas</div><div class="p-stat-val" style="color:var(--corp-green)">${supportFacturaCount}</div></div>
              <div class="p-stat"><div class="p-stat-label">GLPI</div><div class="p-stat-val" style="color:var(--corp-amber)">${supportGlpiCount}</div></div>
              <div class="p-stat"><div class="p-stat-label">Pedidos</div><div class="p-stat-val" style="color:${c}">${supportPedidoCount}</div></div>
            </div>`
          : `<div class="persona-stats">
              <div class="p-stat"><div class="p-stat-label">Total</div><div class="p-stat-val" style="color:${c};font-size:11px">${abr(totalCOP)}</div></div>
              <div class="p-stat"><div class="p-stat-label">Registros</div><div class="p-stat-val">${supportRows.length}</div></div>
              <div class="p-stat"><div class="p-stat-label">Ganadas</div><div class="p-stat-val" style="color:var(--corp-green)">${totalGanadas}</div></div>
              <div class="p-stat"><div class="p-stat-label">Soporta</div><div class="p-stat-val">${supportedCount}</div></div>
            </div>`)
        : `<div style="font-size:9px;color:var(--text3);font-family:var(--font-display);margin-top:8px;padding:5px 8px;background:rgba(255,255,255,.03);border-radius:6px;letter-spacing:.5px">Sin registros aun</div>`}
    </div>`;
  }).join('');

  if(mode === 'pendientes') {
    const pendingType = selPendiente ? selPendiente.value : '';
    const pendingCategory = selCategoria ? selCategoria.value : '';
    const summaryData = allPending.filter(r => namesMatch(getSalesPendingSupportName(r), selectedSupport));
    let data = summaryData.slice();
    if(pendingType) data = data.filter(r => getSalesPendingType(r) === pendingType);
    if(pendingCategory) data = data.filter(r => getSalesPendingCategory(r) === pendingCategory);
    if(SALES_PENDING_VALUE_ONLY) data = data.filter(r => getSalesPendingInvoiceValue(r) > 0);
    data = data.sort((a,b)=>{
      const byValue = getSalesPendingInvoiceValue(b) - getSalesPendingInvoiceValue(a);
      if(byValue !== 0) return byValue;
      return getRowClientName(a).localeCompare(getRowClientName(b), 'es');
    });

    const totalFacturas = data.reduce((sum,row)=>sum+getSalesPendingInvoiceValue(row),0);
    const totalRecords = summaryData.length;
    const totalPedidos = summaryData.filter(r=>normalizeCategoryValue(getSalesPendingType(r)).includes('pedido')).length;
    const totalGlpi = summaryData.filter(r=>normalizeCategoryValue(getSalesPendingType(r)).includes('glpi')).length;
    const totalRemisiones = summaryData.filter(r=>normalizeCategoryValue(getSalesPendingType(r)).includes('remision')).length;
    const categoryNames = [...new Set(['Compra','Garantias','Factura', ...data.map(getSalesPendingCategory).filter(Boolean)])];
    const categoryData = categoryNames.map(name=>({
      name,
      val: data.filter(r=>getSalesPendingCategory(r)===name).length
    })).filter(item=>item.val>0);

    host.innerHTML = `
      <div class="section-hd" style="margin-top:16px"><h2>${escHtml(selectedSupport)}</h2><span class="section-tag">PENDIENTES</span></div>
      <div class="kpi-grid kpi-grid-3" style="margin-bottom:16px">
        <div class="kpi exec-card ${pendingType === 'Pedido' ? 'selected' : ''}" style="--ac:var(--corp-cyan)" onclick="${escAttr(jsCall('setSalesPendingFilter', 'pendiente', 'Pedido'))}"><div class="kpi-accent"></div>
          <div class="kpi-label">Pedidos / Compras</div>
          <div class="kpi-val">${totalPedidos}</div>
          <div class="kpi-sub">De ${totalRecords} pendientes</div>
        </div>
        <div class="kpi exec-card ${pendingType === 'GLPI' ? 'selected' : ''}" style="--ac:var(--corp-amber)" onclick="${escAttr(jsCall('setSalesPendingFilter', 'pendiente', 'GLPI'))}"><div class="kpi-accent"></div>
          <div class="kpi-label">GLPI / Garantias</div>
          <div class="kpi-val">${totalGlpi}</div>
          <div class="kpi-sub">De ${totalRecords} pendientes</div>
        </div>
        <div class="kpi exec-card ${pendingType === 'Remision' ? 'selected' : ''}" style="--ac:var(--corp-green)" onclick="${escAttr(jsCall('setSalesPendingFilter', 'pendiente', 'Remision'))}"><div class="kpi-accent"></div>
          <div class="kpi-label">Remisiones / Facturacion</div>
          <div class="kpi-val">${totalRemisiones}</div>
          <div class="kpi-sub">De ${totalRecords} pendientes</div>
        </div>
      </div>

      <div class="g2">
        <div class="chart-card">
          <div class="chart-hd">Categoria</div>
          <div class="donut-wrap">
            <svg id="donut-sales-pending-category" viewBox="0 0 100 100" style="width:130px;height:130px;flex-shrink:0"></svg>
            <div class="donut-leg" id="leg-sales-pending-category"></div>
          </div>
        </div>
        <div class="kpi sales-pending-total-card" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
          <div class="kpi-label">Valor facturas</div>
          <div class="kpi-val">${abr(totalFacturas)}</div>
          <div class="kpi-sub">${fmtCOP(totalFacturas)}</div>
        </div>
      </div>

      <div class="chart-card g1">
        <div class="chart-hd">Detalle pendientes</div>
        <div class="director-table-toolbar">
          <div>
            <div class="director-table-toolbar-label">Hoja PENDIENTES</div>
            <div class="director-table-toolbar-meta">Estos registros se leen desde la hoja PENDIENTES del Excel Sales Support y se mantienen separados del reporte normal.</div>
          </div>
        </div>
        <div class="tbl-wrap">
          ${buildSalesPendingTable(data)}
        </div>
      </div>
    `;

    renderDonut('donut-sales-pending-category', 'leg-sales-pending-category', categoryData,{
      hasActive: Boolean(pendingCategory),
      getOnClick: item => jsCall('toggleSalesPendingFilter', 'categoria', item.name),
      getIsSelected: item => pendingCategory === item.name,
      clickTitle: 'Filtrar por categoria',
      tooltipPrefix: 'Filtrar categoria: '
    });
    return;
  }

  const mes = selMes.value;
  const estado = selEstado.value;
  let data = allSales.filter(r => namesMatch(getSalesSupportName(r), selectedSupport));
  if(mes) data = data.filter(r => getMonth(getRowDateValue(r)) === mes);
  if(estado) data = data.filter(r => cleanDisplayText(r['ESTADO'],'').toUpperCase() === estado);
  const salesBaseData = data.slice();
  const salesFilters = getVisualCrossfilters('sales');
  if(salesFilters.soporta) data = data.filter(r => getSalesSoportaName(r) === salesFilters.soporta);
  data = data.sort((a,b)=>toCOP(b)-toCOP(a));

  const totalCOP = data.reduce((sum,row)=>sum+toCOP(row),0);
  const totalUSD = data.filter(r=>cleanDisplayText(r['MONEDA 2'],'COP').trim().toUpperCase()==='USD').reduce((sum,row)=>sum+(parseMonto(row['MONTO VENTA CLIENTE'])||0),0);
  const utilidadCOP = sumUtilidad(data,'COP');
  const utilidadUSD = sumUtilidad(data,'USD');
  const totalRecords = data.length;
  const totalGanadas = data.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()==='GANADA').length;
  const supportedNames = [...new Set(data.map(r=>getSalesSoportaName(r)).filter(Boolean))];
  const topSoporta = supportedNames.map(name=>({
    name,
    val: data.filter(r=>getSalesSoportaName(r)===name).reduce((sum,row)=>sum+toCOP(row),0)
  })).sort((a,b)=>b.val-a.val).slice(0, TOP_BAR_LIMIT);
  const estadoData = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'].map(name=>({
    name,
    val: data.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()===name).length
  }));

  host.innerHTML = `
    <div class="section-hd" style="margin-top:16px"><h2>${escHtml(selectedSupport)}</h2><span class="section-tag">SALES SUPPORT</span></div>
    ${buildVisualCrossfilterBar('sales', salesBaseData.length, data.length, [['soporta','Soporta']])}
    <div class="kpi-grid kpi-grid-6" style="margin-bottom:16px">
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        ${kpiUSD(totalUSD)}
        <div class="kpi-sub">TRM dia: ${fmtTRMDisplay(getTRM())}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad COP</div>
        <div class="kpi-val">${abr(utilidadCOP)}</div>
        <div class="kpi-sub">${fmtCOP(utilidadCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad USD</div>
        ${kpiUSD(utilidadUSD)}
        <div class="kpi-sub">Liq: ${fmtCOP(utilidadUSD * getTRM())}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ganadas</div>
        <div class="kpi-val">${totalGanadas}</div>
        <div class="kpi-sub">${totalRecords} registros</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Soporta</div>
        <div class="kpi-val">${supportedNames.length}</div>
        <div class="kpi-sub">Comerciales o directores</div>
      </div>
    </div>

    <div class="g2">
      <div class="chart-card">
        <div class="chart-hd">Top soporta</div>
        <div class="bar-list" id="bar-sales-soporta"></div>
      </div>
      <div class="chart-card">
        <div class="chart-hd">Estado registros</div>
        <div class="donut-wrap">
          <svg id="donut-sales-est" viewBox="0 0 100 100" style="width:130px;height:130px;flex-shrink:0"></svg>
          <div class="donut-leg" id="leg-sales-est"></div>
        </div>
      </div>
    </div>

    <div class="chart-card g1">
      <div class="chart-hd">Detalle Sales Support</div>
      <div class="director-table-toolbar">
        <div>
          <div class="director-table-toolbar-label">Reporte separado del forecast</div>
          <div class="director-table-toolbar-meta">Estos registros se leen desde archivos Sales Support y no suman a comerciales ni a directores.</div>
        </div>
      </div>
      <div class="tbl-wrap">
        ${buildSalesTable(data, { clickable:true, sourcePage:'sales' })}
      </div>
    </div>
  `;

  renderBars('bar-sales-soporta', topSoporta, COLORS,null,{
    getOnClick: item => jsCall('setVisualCrossfilter', 'sales', 'soporta', item.name),
    getIsSelected: item => getVisualCrossfilters('sales').soporta === item.name,
    clickTitle: 'Filtrar por soporta',
    tooltipPrefix: 'Filtrar soporta: '
  });
  renderDonut('donut-sales-est', 'leg-sales-est', estadoData,{
    hasActive: Boolean(estado),
    getOnClick: item => jsCall('toggleSalesEstadoFilter', item.name),
    getIsSelected: item => estado === item.name,
    clickTitle: 'Filtrar por estado',
    tooltipPrefix: 'Filtrar estado: '
  });
}

function selectSalesSupport(name){
  const sel = document.getElementById('sel-sales-support');
  if(sel) sel.value = name;
  renderSales();
}

/* ══════════════════════════════════════
   PREVENTA
══════════════════════════════════════ */
function renderPreventa(){
  const allPreventa = getVisiblePreventaData();
  const host = document.getElementById('preventa-content');
  const grid = document.getElementById('preventa-grid');
  const selPreventa = document.getElementById('sel-preventa');
  const selMes = document.getElementById('sel-preventa-mes');
  const selEstado = document.getElementById('sel-preventa-estado');
  if(!host || !grid || !selPreventa || !selMes || !selEstado) return;

  const namesFromData = [...new Set(allPreventa.map(r=>getPreventaName(r)).filter(Boolean))];
  const namesFromFiles = (LOADED_PREVENTA_FILES || []).map(f=>f.name).filter(Boolean);
  const allNames = [...new Set([...namesFromData, ...namesFromFiles])].sort((a,b)=>a.localeCompare(b,'es'));

  const current = selPreventa.value;
  selPreventa.innerHTML = buildOptionList(allNames);
  if(current && allNames.includes(current)) selPreventa.value = current;
  else if(allNames[0]) selPreventa.value = allNames[0];

  if(!allNames.length){
    grid.innerHTML = '';
    host.innerHTML = `<div class="chart-card g1"><div style="font-size:12px;color:var(--text2)">No hay archivos cargados en Grupo preventa.</div></div>`;
    return;
  }

  const selected = selPreventa.value;
  const selectedBaseRows = selected ? allPreventa.filter(r => namesMatch(getPreventaName(r), selected)) : allPreventa;
  const mes = syncMonthSelectOptions('sel-preventa-mes', getForecastMonths(selectedBaseRows));
  const estado = selEstado.value;

  grid.innerHTML = allNames.map((name, idx)=>{
    const rows = allPreventa.filter(r => namesMatch(getPreventaName(r), name));
    const totalCOP = rows.reduce((sum,row)=>sum+toCOP(row),0);
    const ganadas = rows.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()==='GANADA').length;
    const pendientes = rows.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()==='PENDIENTE').length;
    const c = COLORS[idx % COLORS.length];
    const hasData = rows.length > 0;
    const active = namesMatch(selected, name) ? 'selected' : '';
    return `<div class="persona-card ${active} ${hasData?'':'no-data'}" onclick="${escAttr(jsCall('selectPreventa', name))}">
      <div class="persona-avatar" style="background:${c}${hasData?'25':'10'};border:2px solid ${c}${hasData?'50':'20'};color:${hasData?c:'var(--text2)'}">${escHtml(initials(name))}</div>
      <div class="persona-name" style="color:${hasData?'var(--text)':'var(--text2)'}">${escHtml(name)}</div>
      <div class="persona-role">Preventa</div>
      ${hasData
        ? `<div class="persona-stats">
            <div class="p-stat"><div class="p-stat-label">Total</div><div class="p-stat-val" style="color:${c};font-size:11px">${abr(totalCOP)}</div></div>
            <div class="p-stat"><div class="p-stat-label">Registros</div><div class="p-stat-val">${rows.length}</div></div>
            <div class="p-stat"><div class="p-stat-label">Ganadas</div><div class="p-stat-val" style="color:var(--corp-green)">${ganadas}</div></div>
            <div class="p-stat"><div class="p-stat-label">Pend.</div><div class="p-stat-val" style="color:var(--corp-amber)">${pendientes}</div></div>
          </div>`
        : `<div style="font-size:9px;color:var(--text3);font-family:var(--font-display);margin-top:8px;padding:5px 8px;background:rgba(255,255,255,.03);border-radius:6px;letter-spacing:.5px">Sin registros aun</div>`}
    </div>`;
  }).join('');

  let data = selectedBaseRows.slice();
  if(mes) data = data.filter(r => getMonth(getRowDateValue(r)) === mes);
  if(estado) data = data.filter(r => cleanDisplayText(r['ESTADO'],'').toUpperCase() === estado);
  const preventaBaseData = data.slice();
  const preventaFilters = getVisualCrossfilters('preventa');
  if(preventaFilters.linea) data = data.filter(r => getRowLineName(r) === preventaFilters.linea);
  data = data.sort((a,b)=>toCOP(b)-toCOP(a));

  const totalCOP = data.reduce((sum,row)=>sum+toCOP(row),0);
  const totalUSD = data.filter(r=>cleanDisplayText(r['MONEDA 2'],'COP').trim().toUpperCase()==='USD').reduce((sum,row)=>sum+(parseMonto(row['MONTO VENTA CLIENTE'])||0),0);
  const utilidadCOP = sumUtilidad(data,'COP');
  const utilidadUSD = sumUtilidad(data,'USD');
  const ganadas = data.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()==='GANADA');
  const lineData = buildLineValueData(data);
  const estadoData = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'].map(name=>({
    name,
    val: data.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()===name).length
  }));

  host.innerHTML = `
    <div class="section-hd" style="margin-top:16px"><h2>${escHtml(selected)}</h2><span class="section-tag">PREVENTA</span></div>
    ${buildVisualCrossfilterBar('preventa', preventaBaseData.length, data.length, [['linea','Linea']])}
    <div class="kpi-grid kpi-grid-6" style="margin-bottom:16px">
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        ${kpiUSD(totalUSD)}
        <div class="kpi-sub">TRM dia: ${fmtTRMDisplay(getTRM())}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad COP</div>
        <div class="kpi-val">${abr(utilidadCOP)}</div>
        <div class="kpi-sub">${fmtCOP(utilidadCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad USD</div>
        ${kpiUSD(utilidadUSD)}
        <div class="kpi-sub">Liq: ${fmtCOP(utilidadUSD * getTRM())}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ganadas</div>
        <div class="kpi-val">${ganadas.length}</div>
        <div class="kpi-sub">${data.length} registros</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Registros</div>
        <div class="kpi-val">${data.length}</div>
        <div class="kpi-sub">No suma al forecast</div>
      </div>
    </div>

    <div class="g2">
      <div class="chart-card">
        <div class="chart-hd">Top lineas</div>
        <div class="bar-list" id="bar-preventa-lineas"></div>
      </div>
      <div class="chart-card">
        <div class="chart-hd">Estado registros</div>
        <div class="donut-wrap">
          <svg id="donut-preventa-est" viewBox="0 0 100 100" style="width:130px;height:130px;flex-shrink:0"></svg>
          <div class="donut-leg" id="leg-preventa-est"></div>
        </div>
      </div>
    </div>

    <div class="chart-card g1">
      <div class="chart-hd">Detalle Preventa</div>
      <div class="director-table-toolbar">
        <div>
          <div class="director-table-toolbar-label">Reporte separado del forecast comercial</div>
          <div class="director-table-toolbar-meta">Estos registros se leen desde Grupo preventa y no suman a gerencia, directores, ejecutivos, divisas, marcas ni resumen.</div>
        </div>
      </div>
      <div class="tbl-wrap">
        ${buildTable(data, { clickable:true, sourcePage:'preventa' })}
      </div>
    </div>
  `;

  renderBars('bar-preventa-lineas', lineData, COLORS,null,{
    getOnClick: item => jsCall('setVisualCrossfilter', 'preventa', 'linea', item.name),
    getIsSelected: item => getVisualCrossfilters('preventa').linea === item.name,
    clickTitle: 'Filtrar por linea',
    tooltipPrefix: 'Filtrar linea: '
  });
  renderDonut('donut-preventa-est', 'leg-preventa-est', estadoData,{
    hasActive: Boolean(estado),
    getOnClick: item => jsCall('togglePreventaEstadoFilter', item.name),
    getIsSelected: item => estado === item.name,
    clickTitle: 'Filtrar por estado',
    tooltipPrefix: 'Filtrar estado: '
  });
}

function selectPreventa(name){
  const sel = document.getElementById('sel-preventa');
  if(sel) sel.value = name;
  renderPreventa();
}

/* ══════════════════════════════════════
   DIVISAS
══════════════════════════════════════ */
function renderDivisas(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const trm=getTRM();
  const estadoDetalleEl = document.getElementById('sel-divisa-estado');
  const estadoDetalle = estadoDetalleEl ? estadoDetalleEl.value : 'GANADA';
  
  const getDivisaMoneda = row => cleanDisplayText(row['MONEDA 2'], '').trim().toUpperCase();
  const usdData=ALL_DATA.filter(r=>getDivisaMoneda(r)==='USD');
  const copData=ALL_DATA.filter(r=>getDivisaMoneda(r)==='COP');
  const totalDivisaData=[...copData, ...usdData];
  const usdDetailData=(estadoDetalle ? usdData.filter(r=>r['ESTADO']===estadoDetalle) : usdData).sort((a,b)=>toCOP(b)-toCOP(a));
  const copDetailData=(estadoDetalle ? copData.filter(r=>r['ESTADO']===estadoDetalle) : copData).sort((a,b)=>toCOP(b)-toCOP(a));
  const usdVisible = DIVISAS_DETAIL_LIMITS.USD || 10;
  const copVisible = DIVISAS_DETAIL_LIMITS.COP || 10;
  const usdRemaining = Math.max(usdDetailData.length - usdVisible, 0);
  const copRemaining = Math.max(copDetailData.length - copVisible, 0);
  
  const totalUSD=usdData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const totalCOP=copData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const copGanadas = copData.filter(r=>r['ESTADO']==='GANADA').length;
  const usdGanadas = usdData.filter(r=>r['ESTADO']==='GANADA').length;
  const copPendientes = copData.filter(r=>r['ESTADO']==='PENDIENTE').length;
  const usdPendientes = usdData.filter(r=>r['ESTADO']==='PENDIENTE').length;
  const copPerdidas = copData.filter(r=>r['ESTADO']==='PERDIDA').length;
  const usdPerdidas = usdData.filter(r=>r['ESTADO']==='PERDIDA').length;
  const copAplazadas = copData.filter(r=>r['ESTADO']==='APLAZADO').length;
  const usdAplazadas = usdData.filter(r=>r['ESTADO']==='APLAZADO').length;
  const copGanadasPct = copData.length ? copGanadas / copData.length : 0;
  const usdGanadasPct = usdData.length ? usdGanadas / usdData.length : 0;
  const usdLiqCOP=totalUSD*trm;
  const granTotal=totalCOP+usdLiqCOP;
  const totalStatusRows = (estado) => totalDivisaData.filter(r=>normalizeEstado(r['ESTADO']) === estado);
  const totalStatusValue = (rows) => rows.reduce((sum,row)=>sum+toCOP(row),0);
  const totalGanadasRows = totalStatusRows('GANADA');
  const totalPendientesRows = totalStatusRows('PENDIENTE');
  const totalAplazadasRows = totalStatusRows('APLAZADO');
  const totalPerdidasRows = totalStatusRows('PERDIDA');

  const divisasTotalCards = document.getElementById('divisas-total-cards');
  if(divisasTotalCards) {
    divisasTotalCards.innerHTML = `
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total negocios</div>
        <div class="kpi-val">${fmtNum(totalDivisaData.length)}</div>
        <div class="kpi-sub">${fmtCOP(granTotal)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total ganadas</div>
        <div class="kpi-val">${fmtNum(totalGanadasRows.length)}</div>
        <div class="kpi-sub">${fmtCOP(totalStatusValue(totalGanadasRows))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total pendientes</div>
        <div class="kpi-val">${fmtNum(totalPendientesRows.length)}</div>
        <div class="kpi-sub">${fmtCOP(totalStatusValue(totalPendientesRows))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-red)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total aplazadas</div>
        <div class="kpi-val">${fmtNum(totalAplazadasRows.length)}</div>
        <div class="kpi-sub">${fmtCOP(totalStatusValue(totalAplazadasRows))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total perdidas</div>
        <div class="kpi-val">${fmtNum(totalPerdidasRows.length)}</div>
        <div class="kpi-sub">${fmtCOP(totalStatusValue(totalPerdidasRows))}</div>
      </div>
    `;
  }
  
  document.getElementById('divisas-cards').innerHTML=`
    <div class="divisa-card cop">
      <div class="divisa-label" style="color:var(--cop-color)"><span class="flag">🇨🇴</span> COP — Peso Colombiano</div>
      <div class="divisa-main" style="color:var(--cop-color)">${abr(totalCOP)}</div>
      <div class="divisa-sub">${fmtCOP(totalCOP)}</div>
      <div class="divisa-stats">
        <div><div class="d-stat-label">Negocios</div><div class="d-stat-val" style="color:var(--cop-color)">${copData.length}</div></div>
        <div><div class="d-stat-label">Ganadas</div><div class="d-stat-val" style="color:var(--corp-green)">${copGanadas}</div></div>
        <div><div class="d-stat-label">Pendientes</div><div class="d-stat-val" style="color:var(--corp-amber)">${copPendientes}</div></div>
        <div><div class="d-stat-label">Perdidas</div><div class="d-stat-val" style="color:var(--corp-blue2)">${copPerdidas}</div></div>
        <div><div class="d-stat-label">Aplazadas</div><div class="d-stat-val" style="color:var(--corp-red)">${copAplazadas}</div></div>
        <div><div class="d-stat-label">% Ganadas</div><div class="d-stat-val" style="color:var(--cop-color)">${fmtPct(copGanadasPct)}</div></div>
      </div>
    </div>
    <div class="divisa-card usd">
      <div class="divisa-label" style="color:var(--usd-color)"><span class="flag">🇺🇸</span> USD — Dólar Americano</div>
      <div class="divisa-main" style="color:var(--usd-color)">${fmtUSD(totalUSD)}</div>
      <div class="divisa-sub">TRM ${fmtTRMDisplay(trm)} → Liquidado ${abr(usdLiqCOP)}</div>
      <div class="divisa-stats">
        <div><div class="d-stat-label">Negocios</div><div class="d-stat-val" style="color:var(--usd-color)">${usdData.length}</div></div>
        <div><div class="d-stat-label">Ganadas</div><div class="d-stat-val" style="color:var(--corp-green)">${usdGanadas}</div></div>
        <div><div class="d-stat-label">Pendientes</div><div class="d-stat-val" style="color:var(--corp-amber)">${usdPendientes}</div></div>
        <div><div class="d-stat-label">Perdidas</div><div class="d-stat-val" style="color:var(--corp-blue2)">${usdPerdidas}</div></div>
        <div><div class="d-stat-label">Aplazadas</div><div class="d-stat-val" style="color:var(--corp-red)">${usdAplazadas}</div></div>
        <div><div class="d-stat-label">% Ganadas</div><div class="d-stat-val" style="color:var(--usd-color)">${fmtPct(usdGanadasPct)}</div></div>
      </div>
    </div>
  `;
  
  // Table USD detail
  document.getElementById('tbl-usd').innerHTML=`<table class="responsive-table">
    <thead><tr><th>Ejecutivo</th><th>Cliente</th><th>Producto</th><th>USD</th><th>COP Liquidado</th><th>Estado</th></tr></thead>
    <tbody>${usdDetailData.length ? usdDetailData.slice(0, usdVisible).map(r=>{
      const usd=parseMonto(r['MONTO VENTA CLIENTE'])||0;
      const liq=usd*trm;
      const estado = cleanDisplayText(r['ESTADO'],'—');
      const estadoClass = getEstadoBadgeClass(estado);
      const rowAttrs = r.__RID
        ? ` class="table-row-action" onclick="${escAttr(jsCall('openNegocioDetailById', r.__RID, 'divisas'))}" title="Abrir detalle del negocio"`
        : '';
      return `<tr${rowAttrs}>
        <td data-label="Ejecutivo">${escHtml((r['COMERCIAL']||'').split(' ')[0])}</td>
        <td data-label="Cliente">${escHtml(r['CLIENTE']||'—')}</td>
        <td style="max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" data-label="Producto">${escHtml(r['PRODUCTO']||'—')}</td>
        <td class="td-mono td-usd" data-label="USD">${fmtUSD(usd)}</td>
        <td class="td-mono td-cop" data-label="COP Liquidado">${fmtCOP(liq)}</td>
        <td data-label="Estado"><span class="badge badge-${estadoClass}">${escHtml(estado)}</span></td>
      </tr>`;
    }).join('') : `<tr><td colspan="6" style="text-align:center;color:var(--text2);padding:20px 14px">Sin negocios ${estadoDetalle ? estadoDetalle.toLowerCase() : ''} en USD.</td></tr>`}</tbody>
  </table>${usdRemaining > 0 ? `<div class="table-more-wrap"><button type="button" class="table-more-btn" onclick="showMoreDivisaRows('USD')">Ver mas (${usdRemaining})</button></div>` : ''}`;
  
  // Table COP detail
  document.getElementById('tbl-cop').innerHTML=`<table class="responsive-table">
    <thead><tr><th>Ejecutivo</th><th>Cliente</th><th>Producto</th><th>COP</th><th>Estado</th></tr></thead>
    <tbody>${copDetailData.length ? copDetailData.slice(0, copVisible).map(r=>{
      const cop=parseMonto(r['MONTO VENTA CLIENTE'])||0;
      const estado = cleanDisplayText(r['ESTADO'],'—');
      const estadoClass = getEstadoBadgeClass(estado);
      const rowAttrs = r.__RID
        ? ` class="table-row-action" onclick="${escAttr(jsCall('openNegocioDetailById', r.__RID, 'divisas'))}" title="Abrir detalle del negocio"`
        : '';
      return `<tr${rowAttrs}>
        <td data-label="Ejecutivo">${escHtml((r['COMERCIAL']||'').split(' ')[0])}</td>
        <td data-label="Cliente">${escHtml(r['CLIENTE']||'—')}</td>
        <td style="max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" data-label="Producto">${escHtml(r['PRODUCTO']||'—')}</td>
        <td class="td-mono td-cop" data-label="COP">${fmtCOP(cop)}</td>
        <td data-label="Estado"><span class="badge badge-${estadoClass}">${escHtml(estado)}</span></td>
      </tr>`;
    }).join('') : `<tr><td colspan="5" style="text-align:center;color:var(--text2);padding:20px 14px">Sin negocios ${estadoDetalle ? estadoDetalle.toLowerCase() : ''} en COP.</td></tr>`}</tbody>
  </table>${copRemaining > 0 ? `<div class="table-more-wrap"><button type="button" class="table-more-btn" onclick="showMoreDivisaRows('COP')">Ver mas (${copRemaining})</button></div>` : ''}`;
  
  // Tabla resumen consolidado
  const dirs=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))];
  document.getElementById('tbl-resumen-divisas').innerHTML=`<table class="responsive-table">
    <thead><tr><th>Director</th><th>Negos COP</th><th>Valor COP</th><th>Negos USD</th><th>Valor USD</th><th>Liq. USD→COP</th><th>TOTAL COP</th></tr></thead>
    <tbody>${dirs.map(d=>{
      const dd=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d);
      const dc=dd.filter(r=>(r['MONEDA 2']||'').trim()==='COP');
      const du=dd.filter(r=>(r['MONEDA 2']||'').trim()==='USD');
      const vCOP=dc.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const vUSD=du.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const liq=vUSD*trm;
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:700;color:var(--text)" data-label="Director">${escHtml(d)}</td>
        <td class="td-mono" data-label="Negos COP">${dc.length}</td>
        <td class="td-mono td-cop" data-label="Valor COP">${fmtCOP(vCOP)}</td>
        <td class="td-mono" data-label="Negos USD">${du.length}</td>
        <td class="td-mono td-usd" data-label="Valor USD">${fmtUSD(vUSD)}</td>
        <td class="td-mono td-usd" data-label="Liq. USD→COP">${fmtCOP(liq)}</td>
        <td class="td-mono" style="color:var(--text);font-weight:600" data-label="TOTAL COP">${fmtCOP(vCOP+liq)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td style="font-family:var(--font-display);font-weight:800;color:var(--text)" data-label="Director">TOTAL GENERAL</td>
      <td class="td-mono" data-label="Negos COP">${copData.length}</td>
      <td class="td-mono td-cop" style="font-weight:700" data-label="Valor COP">${fmtCOP(totalCOP)}</td>
      <td class="td-mono" data-label="Negos USD">${usdData.length}</td>
      <td class="td-mono td-usd" style="font-weight:700" data-label="Valor USD">${fmtUSD(totalUSD)}</td>
      <td class="td-mono td-usd" style="font-weight:700" data-label="Liq. USD→COP">${fmtCOP(usdLiqCOP)}</td>
      <td class="td-mono" style="color:var(--corp-cyan);font-weight:800;font-size:13px" data-label="TOTAL COP">${fmtCOP(granTotal)}</td>
    </tr></tbody>
  </table>`;
}

/* ══════════════════════════════════════
   MARCAS
══════════════════════════════════════ */
function renderMarcas(){
  const BASE_DATA = getVisibleMarcasData();
  const marcaFilters = getVisualCrossfilters('marcas');
  let ALL_DATA = BASE_DATA.slice();
  if(marcaFilters.marca) ALL_DATA = ALL_DATA.filter(r => getRowBrandName(r) === marcaFilters.marca);
  if(marcaFilters.linea) ALL_DATA = ALL_DATA.filter(r => getRowLineName(r) === marcaFilters.linea);
  const pageTag = document.querySelector('#page-marcas .section-tag');
  if(pageTag) pageTag.textContent = isOscarMarcasGlobalScope() ? 'TODOS LOS USUARIOS' : 'TOP POR CATEGORIA';
  const marcasCrossfilters = document.getElementById('marcas-crossfilters');
  if(marcasCrossfilters) {
    marcasCrossfilters.innerHTML = buildVisualCrossfilterBar('marcas', BASE_DATA.length, ALL_DATA.length, [['marca','Marca'],['linea','Linea']]);
  }
  if(!ALL_DATA.length){
    const empty = `<div style="padding:24px 16px;text-align:center;color:var(--text2);font-family:var(--font-body)">Sin registros disponibles para esta vista.</div>`;
    const barMarcas = document.getElementById('bar-marcas');
    const barMarcasMore = document.getElementById('bar-marcas-more');
    const barLineas = document.getElementById('bar-lineas');
    const tblMarcaEj = document.getElementById('tbl-marca-ej');
    const legMarca = document.getElementById('leg-marca');
    const legLinea = document.getElementById('leg-linea2');
    const donutMarca = document.getElementById('donut-marca');
    const donutLinea = document.getElementById('donut-linea2');
    if(barMarcas) barMarcas.innerHTML = empty;
    if(barMarcasMore) barMarcasMore.innerHTML = '';
    if(barLineas) barLineas.innerHTML = empty;
    if(tblMarcaEj) tblMarcaEj.innerHTML = empty;
    if(legMarca) legMarca.innerHTML = '';
    if(legLinea) legLinea.innerHTML = '';
    if(donutMarca) donutMarca.innerHTML = '';
    if(donutLinea) donutLinea.innerHTML = '';
    return;
  }
  
  const marcaData=buildBrandValueData(ALL_DATA);
  const marcas=marcaData.map(item => item.name);
  const visibleMarcaData = marcaData.slice(0, Math.max(MARCAS_BAR_LIMIT || MARCAS_BAR_INITIAL, MARCAS_BAR_INITIAL));
  renderBars('bar-marcas',visibleMarcaData,COLORS,null,{
    getOnClick: item => jsCall('openMarcaLineaDetail', 'marca', item.name, 'marcas'),
    clickTitle: 'Abrir detalle de marca',
    nameClass: 'w160'
  });
  const barMarcasMore = document.getElementById('bar-marcas-more');
  const marcasRemaining = Math.max(marcaData.length - visibleMarcaData.length, 0);
  if(barMarcasMore) {
    barMarcasMore.innerHTML = marcasRemaining > 0
      ? `<button type="button" class="table-more-btn" onclick="showMoreMarcasBars()">Ver mas (${marcasRemaining})</button>`
      : '';
  }
  renderDonut('donut-marca','leg-marca',marcaData,{
    hasActive: Boolean(marcaFilters.marca),
    getOnClick: item => jsCall('setVisualCrossfilter', 'marcas', 'marca', item.name),
    getIsSelected: item => getVisualCrossfilters('marcas').marca === item.name,
    clickTitle: 'Filtrar por marca',
    tooltipPrefix: 'Filtrar marca: '
  });
  
  const linData=buildLineValueData(ALL_DATA);
  renderBars('bar-lineas',linData,COLORS,null,{
    getOnClick: item => jsCall('openMarcaLineaDetail', 'linea', item.name, 'marcas'),
    clickTitle: 'Abrir detalle de linea',
    nameClass: 'w160'
  });
  renderDonut('donut-linea2','leg-linea2',linData,{
    hasActive: Boolean(marcaFilters.linea),
    getOnClick: item => jsCall('setVisualCrossfilter', 'marcas', 'linea', item.name),
    getIsSelected: item => getVisualCrossfilters('marcas').linea === item.name,
    clickTitle: 'Filtrar por linea',
    tooltipPrefix: 'Filtrar linea: '
  });
  
  // Marca por ejecutivo
  const directorSelect = document.getElementById('sel-marca-director');
  const directorFilterWrap = document.getElementById('marca-director-filter');
  const directorOptions = [...new Set([
    ...ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()),
    ...Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim())
  ].filter(Boolean))].sort();
  let selectedDirector = '';
  if(directorSelect){
    const prev = directorSelect.value;
    const opts = directorOptions.length > 1
      ? [optionHtml('', 'Todos los directores', false), buildOptionList(directorOptions)]
      : [buildOptionList(directorOptions)];
    directorSelect.innerHTML = opts.join('');
    if(prev === '' && directorOptions.length > 1) directorSelect.value = '';
    else if(prev && directorOptions.includes(prev)) directorSelect.value = prev;
    else if(directorOptions.length === 1) directorSelect.value = directorOptions[0];
    else if(directorOptions.length > 1) directorSelect.value = '';
    selectedDirector = directorSelect.value || '';
  }
  if(directorFilterWrap) directorFilterWrap.style.display = directorOptions.length > 1 ? '' : 'none';

  const marcaExecData = selectedDirector
    ? ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===selectedDirector)
    : ALL_DATA;
  const execs=[...new Set(marcaExecData.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  document.getElementById('tbl-marca-ej').innerHTML=`<table class="responsive-table">
    <thead><tr><th>Ejecutivo</th><th>Top Marca</th><th>Cantidad</th></tr></thead>
    <tbody>${execs.length ? execs.map(e=>{
      const ed=marcaExecData.filter(r=>r['COMERCIAL']===e);
      const marcaCounts=marcas.map(m=>ed.filter(r=>getRowBrandName(r)===m).length);
      const topIdx=marcaCounts.indexOf(Math.max(...marcaCounts));
      const execDirector = selectedDirector ? cleanDisplayText(selectedDirector, '') : '';
      const topMarca = marcas[topIdx] || '—';
      const topMarcaCount = topIdx >= 0 ? marcaCounts[topIdx] || 0 : 0;
      const topColor = topIdx >= 0 ? COLORS[topIdx%COLORS.length] : 'var(--text3)';
      return `<tr class="table-row-action" onclick="${escAttr(jsCall('openExecNegociosFromMarcas', e, execDirector, topMarca))}" title="Abrir negocios del ejecutivo">
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)" data-label="Ejecutivo">${escHtml(e)}</td>
        <td style="font-family:var(--font-display);font-weight:700;color:${topColor}" data-label="Top Marca">${escHtml(topMarca)}</td>
        <td class="td-mono" data-label="Cantidad">${fmtNum(topMarcaCount)}</td>
      </tr>`;
    }).join('') : `<tr><td colspan="3" style="text-align:center;color:var(--text2);padding:20px 14px">Sin ejecutivos con datos para este grupo.</td></tr>`}</tbody>
  </table>`;
}

/* ══════════════════════════════════════
   RESUMEN TOTAL
══════════════════════════════════════ */
function renderResumen(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const trm=getTRM();
  const months = getForecastMonths(ALL_DATA);
  const monthHeaderCells = months.map(month => `<th>${getMonthShortLabel(month)}</th>`).join('');
  
  const usdData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='USD');
  const copData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='COP');
  const totalUSD=usdData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const totalCOP=copData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const usdLiq=totalUSD*trm;
  
  document.getElementById('kpi-resumen-top').innerHTML=`
    <div class="kpi" style="--ac:var(--cop-color)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total Puro COP</div>
      <div class="kpi-val">${abr(totalCOP)}</div>
      <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      <span class="kpi-badge cop">COP</span>
    </div>
    <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total USD Liquidado COP</div>
      <div class="kpi-val">${abr(usdLiq)}</div>
      <div class="kpi-sub">${fmtUSD(totalUSD)} × ${fmtTRMDisplay(trm)}</div>
      <span class="kpi-badge usd">USD→COP</span>
    </div>
    <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
      <div class="kpi-label">GRAN TOTAL COP</div>
      <div class="kpi-val">${abr(totalCOP+usdLiq)}</div>
      <div class="kpi-sub">${fmtCOP(totalCOP+usdLiq)}</div>
    </div>
  `;
  
  // Tabla solo COP por director/ejecutivo
  const dirs=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))];
  const execs=[...new Set(ALL_DATA.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  
  document.getElementById('tbl-total-cop').innerHTML=`<table class="responsive-table">
    <thead><tr><th>Ejecutivo</th><th>Director</th><th>Negocios COP</th>${monthHeaderCells}<th>Total COP</th></tr></thead>
    <tbody>${execs.map(e=>{
      const ed=copData.filter(r=>r['COMERCIAL']===e);
      const dir=ALL_DATA.find(r=>r['COMERCIAL']===e);
      const monthlyValues = months.map(month => ed.filter(r=>getMonth(getRowDateValue(r))===month).reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0));
      const tot=monthlyValues.reduce((sum, val)=>sum+val,0);
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)" data-label="Ejecutivo">${escHtml(e)}</td>
        <td style="color:var(--text2)" data-label="Director">${escHtml(dir?dir['DIRECTOR']||'—':'—')}</td>
        <td class="td-mono" data-label="Negocios COP">${ed.length}</td>
        ${monthlyValues.map((val,idx)=>`<td class="td-mono td-cop" data-label="${escAttr(getMonthShortLabel(months[idx]))}">${val>0?abr(val):'—'}</td>`).join('')}
        <td class="td-mono td-cop" style="font-weight:700" data-label="Total COP">${fmtCOP(tot)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;color:var(--cop-color)">SUBTOTAL COP</td>
      <td class="td-mono" data-label="Negocios COP">${copData.length}</td>
      ${months.map(month => `<td class="td-mono td-cop" data-label="${escAttr(getMonthShortLabel(month))}">${abr(copData.filter(r=>getMonth(getRowDateValue(r))===month).reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0))}</td>`).join('')}
      <td class="td-mono td-cop" style="font-weight:800;font-size:13px" data-label="Total COP">${fmtCOP(totalCOP)}</td>
    </tr></tbody>
  </table>`;
  
  // Tabla USD detail + liquidado
  document.getElementById('tbl-total-usd').innerHTML=`<table class="responsive-table">
    <thead><tr><th>Ejecutivo</th><th>Director</th><th>Negocios USD</th>${monthHeaderCells}<th>Total USD</th><th>TRM</th><th>Total Liquidado COP</th></tr></thead>
    <tbody>${execs.map(e=>{
      const ed=usdData.filter(r=>r['COMERCIAL']===e);
      if(!ed.length) return '';
      const dir=ALL_DATA.find(r=>r['COMERCIAL']===e);
      const monthlyValues = months.map(month => ed.filter(r=>getMonth(getRowDateValue(r))===month).reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0));
      const tot=monthlyValues.reduce((sum, val)=>sum+val,0);
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)" data-label="Ejecutivo">${escHtml(e)}</td>
        <td style="color:var(--text2)" data-label="Director">${escHtml(dir?dir['DIRECTOR']||'—':'—')}</td>
        <td class="td-mono" data-label="Negocios USD">${ed.length}</td>
        ${monthlyValues.map((val,idx)=>`<td class="td-mono td-usd" data-label="${escAttr(getMonthShortLabel(months[idx]))}">${val>0?fmtUSD(val):'—'}</td>`).join('')}
        <td class="td-mono td-usd" style="font-weight:700" data-label="Total USD">${fmtUSD(tot)}</td>
        <td class="td-mono" style="color:var(--corp-cyan)" data-label="TRM">${fmtTRMDisplay(trm)}</td>
        <td class="td-mono td-cop" style="font-weight:700" data-label="Total Liquidado COP">${fmtCOP(tot*trm)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;color:var(--usd-color)">SUBTOTAL USD</td>
      <td class="td-mono" data-label="Negocios USD">${usdData.length}</td>
      ${months.map(month=>`<td data-label="${escAttr(getMonthShortLabel(month))}"></td>`).join('')}
      <td class="td-mono td-usd" style="font-weight:800" data-label="Total USD">${fmtUSD(totalUSD)}</td>
      <td class="td-mono" style="color:var(--corp-cyan)" data-label="TRM">${fmtTRMDisplay(trm)}</td>
      <td class="td-mono td-usd" style="font-weight:800;font-size:13px" data-label="Total Liquidado COP">${fmtCOP(usdLiq)}</td>
    </tr></tbody>
  </table>`;
  
  // Consolidado final
  document.getElementById('tbl-consolidado').innerHTML=`<table class="responsive-table">
    <thead><tr><th>Director</th><th>Ejecutivo</th><th>Total COP</th><th>Total USD</th><th>USD Liq. COP</th><th>TOTAL CONSOLIDADO</th></tr></thead>
    <tbody>${dirs.flatMap(d=>{
      const dejecs=[...new Set(ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d).map(r=>r['COMERCIAL']))];
      return dejecs.map((e,ei)=>{
        const ed=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d&&r['COMERCIAL']===e);
        const eCOP=ed.filter(r=>(r['MONEDA 2']||'').trim()==='COP').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
        const eUSD=ed.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
        const total=eCOP+eUSD*trm;
        return `<tr>
          <td style="font-family:var(--font-display);font-weight:700;color:var(--text2)" data-label="Director">${ei===0?escHtml(d):''}</td>
          <td style="color:var(--text)" data-label="Ejecutivo">${escHtml(e)}</td>
          <td class="td-mono td-cop" data-label="Total COP">${fmtCOP(eCOP)}</td>
          <td class="td-mono td-usd" data-label="Total USD">${eUSD>0?fmtUSD(eUSD):'—'}</td>
          <td class="td-mono td-usd" data-label="USD Liq. COP">${eUSD>0?fmtCOP(eUSD*trm):'—'}</td>
          <td class="td-mono" style="color:var(--text);font-weight:700" data-label="TOTAL CONSOLIDADO">${fmtCOP(total)}</td>
        </tr>`;
      });
    }).join('')}
    <tr style="border-top:2px solid var(--corp-blue2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;font-size:13px;color:var(--corp-cyan)">GRAN TOTAL</td>
      <td class="td-mono td-cop" style="font-weight:800" data-label="Total COP">${fmtCOP(totalCOP)}</td>
      <td class="td-mono td-usd" style="font-weight:800" data-label="Total USD">${fmtUSD(totalUSD)}</td>
      <td class="td-mono td-usd" style="font-weight:800" data-label="USD Liq. COP">${fmtCOP(usdLiq)}</td>
      <td class="td-mono" style="color:var(--corp-cyan);font-weight:800;font-size:14px" data-label="TOTAL CONSOLIDADO">${fmtCOP(totalCOP+usdLiq)}</td>
    </tr></tbody>
  </table>`;
}

/* ══════════════════════════════════════
   GENERIC TABLE
══════════════════════════════════════ */
function buildTable(data, opts){
  const options = opts || {};
  const sourcePage = options.sourcePage || getActivePageId();
  return `<table class="responsive-table">
    <thead><tr><th>Fecha</th><th>Cliente</th><th>Producto</th><th>Marca</th><th>Línea</th><th>Moneda</th><th>Valor</th><th>COP Total</th><th>Margen</th><th>Estado</th></tr></thead>
    <tbody>${data.length ? data.map(r=>{
      const mon = cleanDisplayText(r['MONEDA 2'], 'COP').trim().toUpperCase();
      const val = parseMonto(r['MONTO VENTA CLIENTE']) || 0;
      const cop = toCOP(r);
      const fecha = cleanDisplayText(formatDateValue(getRowDateValue(r)), 'Sin fecha');
      const cliente = cleanDisplayText(getRowClientName(r), 'Sin cliente');
      const producto = cleanDisplayText(getRowProductName(r), 'Sin proyecto');
      const marca = cleanDisplayText(getRowBrandName(r), 'Sin marca');
      const linea = cleanDisplayText(getRowLineName(r), 'Sin linea');
      const marginRaw = r['MARGEN'];
      const marginText = formatMarginDisplay(marginRaw, '-');
      const estado = cleanDisplayText(r['ESTADO'], 'Sin estado').toUpperCase();
      const estadoClass = getEstadoBadgeClass(estado);
      const rowAttrs = options.clickable && r.__RID
        ? ` class="table-row-action" onclick="${escAttr(jsCall('openNegocioDetailById', r.__RID, sourcePage))}" title="Abrir detalle del negocio"`
        : '';
      return `<tr${rowAttrs}>
        <td class="td-mono" style="font-size:10px" data-label="Fecha">${escHtml(fecha)}</td>
        <td style="color:var(--text)" data-label="Cliente">${escHtml(cliente)}</td>
        <td style="max-width:130px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${escAttr(producto)}" data-label="Producto">${escHtml(producto)}</td>
        <td style="color:var(--corp-cyan)" data-label="Marca">${escHtml(marca)}</td>
        <td style="font-size:10px" data-label="Línea">${escHtml(linea)}</td>
        <td data-label="Moneda"><span class="badge ${mon==='USD'?'badge-PEDIDA':'badge-PENDIENTE'}">${escHtml(mon)}</span></td>
        <td class="td-mono ${mon==='USD'?'td-usd':'td-cop'}" data-label="Valor">${mon==='USD'?fmtUSD(val):fmtCOP(val)}</td>
        <td class="td-mono td-cop" data-label="COP Total">${fmtCOP(cop)}</td>
        <td class="td-mono" style="color:var(--corp-amber)" data-label="Margen">${escHtml(marginText)}</td>
        <td data-label="Estado"><span class="badge badge-${estadoClass}">${escHtml(estado)}</span></td>
      </tr>`;
    }).join('') : `<tr><td colspan="10" style="text-align:center;color:var(--text2)">Sin registros para este filtro.</td></tr>`}</tbody>
  </table>`;
}

/* ══════════════════════════════════════
   DEMO DATA (pre-load for testing)
══════════════════════════════════════ */



async function loadFolderFromSharePoint() {
  // Cargar MSAL
  await new Promise(r => loadMSAL(r));
  if(!initMsalApp()) {
    alert('No se pudo cargar autenticación'); return;
  }
  showLoadingOverlay('Conectando con Microsoft 365...');
  try {
    await spLogin();
  } catch(e) {
    if(typeof isAuthRedirectInProgress === 'function' && isAuthRedirectInProgress(e)) return;
    hideLoadingOverlay();
    alert('Error de autenticación: ' + e.message);
    return;
  }
  updateLoadingStatus('Cargando archivos de SharePoint...');
  try {
    const siteId = await getSiteId();
    const filesToken = await getToken(['Files.Read.All']);
    const { role, directorGroup } = CURRENT_USER;
    ALL_DATA = [];
    SALES_DATA = [];
    SALES_PENDING_DATA = [];
    PREVENTA_DATA = [];
    RECORD_SEQ = 0;
    LOADED_FILES_BY_DIR = {};
    LOADED_SALES_BY_SUPPORT = {};
    LOADED_PREVENTA_FILES = [];

    if(role === 'sales_support') {
      await loadSalesSupportFiles(siteId, filesToken);
    } else if(role === 'ejecutivo') {
      await loadEjecutivoFile(siteId, filesToken);
    } else if(role === 'director') {
      const folder = await getDirectorFolderName(siteId, directorGroup, filesToken);
      await loadDirectorFolder(siteId, folder, filesToken);
    } else {
      // Gerencia — todas las carpetas
      const folders = await getForecastFolders(siteId, filesToken);
      await runWithConcurrencyLimit(folders, FOLDER_LOAD_CONCURRENCY, folderName =>
        loadDirectorFolder(siteId, folderName, filesToken)
      );
      await loadPreventaFolder(siteId, filesToken);
    }
    finalizeLoad();
    hideLoadingOverlay();
    setTimeout(() => {
      syncForecastConnections(siteId).catch(connectionError => {
        console.warn('[FORECAST CONNECTIONS]', connectionError);
      });
    }, 0);
  } catch(e) {
    if(typeof isAuthRedirectInProgress === 'function' && isAuthRedirectInProgress(e)) return;
    hideLoadingOverlay();
    console.error(e);
    alert('Error cargando datos: ' + e.message);
  }
}

let _siteId = null;
let _driveId = null;
let _forecastFolders = null;
let _forecastBasePath = null;
let _siteDrivesById = {};
let _siteIdByPath = {};
const FOLDER_LOAD_CONCURRENCY = 2;
const FILE_LOAD_CONCURRENCY = 3;

function encodeGraphPath(path){
  return String(path || '')
    .split('/')
    .filter(Boolean)
    .map(part => encodeURIComponent(part))
    .join('/');
}

function joinGraphPath(){
  return Array.from(arguments)
    .map(part => String(part || '').trim().replace(/^\/+|\/+$/g, ''))
    .filter(Boolean)
    .join('/');
}

function buildGraphRootUrl(siteId, path, suffix){
  const encodedPath = encodeGraphPath(path);
  const tail = suffix || 'children?$top=50';
  return _driveId
    ? 'https://graph.microsoft.com/v1.0/drives/' + _driveId + '/root:/' + encodedPath + ':/' + tail
    : 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root:/' + encodedPath + ':/' + tail;
}

function getForecastBaseCandidates(){
  const configured = (window.AZURE_CONFIG && AZURE_CONFIG.driveBase) || FORECAST_BASE_FALLBACK;
  const withoutLibrary = String(configured || '')
    .replace(/^Documentos compartidos\//i, '')
    .replace(/^Shared Documents\//i, '')
    .replace(/^Documents\//i, '')
    .replace(/^\/+|\/+$/g, '');
  const withoutComercial = withoutLibrary.replace(/^COMERCIAL\//i, '');
  return [...new Set([
    withoutComercial,
    withoutLibrary,
    'FORECAST 2026',
    FORECAST_BASE_FALLBACK
  ].map(path => path.trim()).filter(Boolean))];
}

async function fetchGraphPathJson(siteId, path, suffix, token){
  const url = buildGraphRootUrl(siteId, path, suffix);
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const d = await r.json().catch(() => ({}));
  return { r, d, url };
}

function graphSearchQueryLiteral(query){
  return encodeURIComponent(String(query || '').replace(/'/g, "''"));
}

function buildDriveRootSearchUrl(driveId, query){
  return 'https://graph.microsoft.com/v1.0/drives/' + driveId
    + "/root/search(q='" + graphSearchQueryLiteral(query) + "')?$top=100";
}

function buildDriveItemUrl(driveId, itemId){
  return 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + itemId;
}

function getSitePathCandidates(){
  const configuredPath = (() => {
    try {
      const url = new URL((window.AZURE_CONFIG && AZURE_CONFIG.siteUrl) || '');
      return url.pathname.replace(/^\/+/, '').replace(/^sites\//i, '').replace(/\/+$/, '');
    } catch(_) {
      return '';
    }
  })();
  return [...new Set([
    'ProvexpressIntranet/comercial',
    configuredPath,
    'ProvexpressIntranet'
  ].filter(Boolean))];
}

async function getSiteIdByPath(path, token){
  if(_siteIdByPath[path]) return _siteIdByPath[path];
  const url = 'https://graph.microsoft.com/v1.0/sites/provexpress.sharepoint.com:/sites/' + encodeGraphPath(path);
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const d = await r.json().catch(() => ({}));
  if(!r.ok || !d.id) throw new Error((d.error && d.error.message) || 'Site no encontrado: ' + path);
  _siteIdByPath[path] = d.id;
  return d.id;
}

async function getSiteDrives(siteId, token){
  if(_siteDrivesById[siteId]) return _siteDrivesById[siteId];
  const r = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/drives', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const d = await r.json().catch(() => ({}));
  if(!r.ok) throw new Error((d.error && d.error.message) || 'No se pudieron consultar documentos de SharePoint');
  _siteDrivesById[siteId] = d.value || [];
  return _siteDrivesById[siteId];
}

async function getSearchDriveCandidates(siteId, token){
  const byDriveId = new Map();
  const addDrive = (site, drive) => {
    if(!drive || !drive.id || byDriveId.has(drive.id)) return;
    byDriveId.set(drive.id, { siteId: site, driveId: drive.id, driveName: drive.name || '', webUrl: drive.webUrl || '' });
  };

  if(_driveId) addDrive(siteId, { id: _driveId, name: 'selected' });

  try {
    (await getSiteDrives(siteId, token)).forEach(drive => addDrive(siteId, drive));
  } catch(e) {
    console.warn('[DRIVE SEARCH] current site drives failed', e);
  }

  for(const sitePath of getSitePathCandidates()) {
    try {
      const candidateSiteId = await getSiteIdByPath(sitePath, token);
      (await getSiteDrives(candidateSiteId, token)).forEach(drive => addDrive(candidateSiteId, drive));
    } catch(e) {
      console.warn('[DRIVE SEARCH] site skipped', sitePath, e);
    }
  }

  return Array.from(byDriveId.values());
}

async function getForecastBasePath(siteId, token){
  if(_forecastBasePath) return _forecastBasePath;
  const authToken = token || await getToken(['Files.Read.All']);
  const attempts = [];
  for(const path of getForecastBaseCandidates()) {
    const { r, d, url } = await fetchGraphPathJson(siteId, path, 'children?$top=1', authToken);
    attempts.push({ path, status: r.status, message: d && d.error && d.error.message });
    if(r.ok && Array.isArray(d.value)) {
      _forecastBasePath = path;
      console.log('[FORECAST BASE]', path);
      return _forecastBasePath;
    }
  }
  console.warn('[FORECAST BASE] no encontrada', attempts);
  throw new Error('No se pudo encontrar la carpeta FORECAST 2026. Rutas probadas: ' + attempts.map(a => a.path + ' (' + a.status + ')').join(', '));
}

async function runWithConcurrencyLimit(items, limit, worker){
  const list = Array.isArray(items) ? items : [];
  const max = Math.max(1, Number(limit) || 1);
  if(!list.length) return [];

  const results = new Array(list.length);
  let cursor = 0;

  async function consume(){
    while(cursor < list.length) {
      const index = cursor++;
      results[index] = await worker(list[index], index);
    }
  }

  const workers = Array.from({ length: Math.min(max, list.length) }, consume);
  await Promise.all(workers);
  return results;
}

async function getForecastConnectionsListId(siteId, token) {
  if(FORECAST_CONNECTIONS_LIST_ID) return FORECAST_CONNECTIONS_LIST_ID;
  const url = 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/lists?$select=id,displayName,name&$top=200';
  const r = await fetch(url, { headers: graphJsonHeaders(token) });
  const d = await r.json();
  if(!r.ok) throw new Error((d.error && d.error.message) || 'No se pudo consultar listas de SharePoint');
  const list = (d.value || []).find(item =>
    item.displayName === FORECAST_CONNECTIONS_LIST_NAME || item.name === FORECAST_CONNECTIONS_LIST_NAME
  );
  if(!list) throw new Error('No se encontro la lista ' + FORECAST_CONNECTIONS_LIST_NAME);
  FORECAST_CONNECTIONS_LIST_ID = list.id;
  return FORECAST_CONNECTIONS_LIST_ID;
}

function buildConnectionMaps(items) {
  const byEmail = {};
  const byName = {};
  const execMap = window.EXECUTIVO_BY_EMAIL || {};
  (items || []).forEach(item => {
    const f = item.fields || {};
    const email = String(f.Title || '').toLowerCase().trim();
    const record = {
      email,
      Nombre: f.Nombre || execMap[email] || '',
      Rol: f.Rol || '',
      Director: f.Director || '',
      UltimaConexion: f.UltimaConexion || '',
      ConteoConexiones: Number(f.ConteoConexiones || 0)
    };
    if(email) byEmail[email] = record;
    const names = [record.Nombre, execMap[email]].filter(Boolean);
    names.forEach(name => { byName[normalizePersonName(name)] = record; });
  });
  FORECAST_CONNECTIONS = { byEmail, byName };
}

async function getForecastConnectionItems(siteId, token, listId) {
  const url = 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/lists/' + listId
    + '/items?$expand=fields($select=Title,Nombre,Rol,Director,UltimaConexion,ConteoConexiones)&$top=500';
  const r = await fetch(url, {
    headers: graphJsonHeaders(token, { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' })
  });
  const d = await r.json();
  if(!r.ok) throw new Error((d.error && d.error.message) || 'No se pudieron leer conexiones');
  return d.value || [];
}

async function syncForecastConnections(siteId) {
  if(!CURRENT_USER || !CURRENT_USER.email) return;
  const token = await getToken(['Sites.ReadWrite.All']);
  const listId = await getForecastConnectionsListId(siteId, token);
  const items = await getForecastConnectionItems(siteId, token, listId);
  const email = String(CURRENT_USER.email || '').toLowerCase().trim();
  const existing = items.find(item => String(item.fields && item.fields.Title || '').toLowerCase().trim() === email);
  const now = new Date().toISOString();
  const fields = {
    Title: email,
    Nombre: CURRENT_USER.name || '',
    Rol: getRoleLabel(CURRENT_USER.role),
    Director: CURRENT_USER.directorGroup || '',
    UltimaConexion: now,
    ConteoConexiones: Number(existing && existing.fields && existing.fields.ConteoConexiones || 0) + 1
  };

  if(existing) {
    const url = 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/lists/' + listId + '/items/' + existing.id + '/fields';
    const r = await fetch(url, { method: 'PATCH', headers: graphJsonHeaders(token), body: JSON.stringify(fields) });
    if(!r.ok) {
      const d = await r.json().catch(() => ({}));
      throw new Error((d.error && d.error.message) || 'No se pudo actualizar conexion');
    }
    existing.fields = { ...existing.fields, ...fields };
  } else {
    const url = 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/lists/' + listId + '/items';
    const r = await fetch(url, { method: 'POST', headers: graphJsonHeaders(token), body: JSON.stringify({ fields }) });
    const d = await r.json().catch(() => ({}));
    if(!r.ok) throw new Error((d.error && d.error.message) || 'No se pudo registrar conexion');
    items.push({ ...d, fields });
  }

  buildConnectionMaps(items);
  console.log('[FORECAST CONNECTIONS]', Object.keys(FORECAST_CONNECTIONS.byEmail).length);
  if(ALL_DATA.length || SALES_DATA.length || SALES_PENDING_DATA.length || PREVENTA_DATA.length) {
    renderVisiblePage();
  }
}

function normalizeFolderName(value){
  return String(value || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/^(grupo|gupo)\s+/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function isPreventaFolderName(value){
  return normalizeFolderName(value) === normalizeFolderName(PREVENTA_FOLDER_NAME);
}

async function getForecastFolders(siteId, token){
  if(_forecastFolders && _forecastFolders.length) return _forecastFolders;
  const authToken = token || await getToken(['Files.Read.All']);
  const forecastBasePath = await getForecastBasePath(siteId, authToken);
  const { d } = await fetchGraphPathJson(siteId, forecastBasePath, 'children?$top=100', authToken);
  const folders = (d.value || [])
    .filter(item => item && item.folder)
    .map(item => item.name)
    .filter(name => /^(Grupo|Gupo)\s+/i.test(name) && !isPreventaFolderName(name));
  if(folders.length) {
    _forecastFolders = folders;
    console.log('[FORECAST FOLDERS]', folders);
    return folders;
  }
  const fallback = ['Grupo Juan David Novoa','Grupo Maria Angelica Caballero','Grupo Oscar Beltran','Gupo Miller Romero'];
  console.warn('[FORECAST FOLDERS] usando fallback', d);
  _forecastFolders = fallback;
  return fallback;
}

async function getDirectorFolderName(siteId, directorGroup, token){
  const folders = await getForecastFolders(siteId, token);
  const target = normalizeFolderName(directorGroup);
  return folders.find(name => normalizeFolderName(name) === target) || (directorGroup.includes('Miller') ? 'Gupo Miller Romero' : 'Grupo ' + directorGroup);
}

async function getPreventaFolderName(siteId, token){
  const authToken = token || await getToken(['Files.Read.All']);
  const forecastBasePath = await getForecastBasePath(siteId, authToken);
  const { d } = await fetchGraphPathJson(siteId, forecastBasePath, 'children?$top=100', authToken);
  const folder = (d.value || [])
    .filter(item => item && item.folder)
    .map(item => item.name)
    .find(name => isPreventaFolderName(name));
  return folder || PREVENTA_FOLDER_NAME;
}

async function getSiteId() {
  if(_siteId) return _siteId;
  const token = await getToken(['Files.Read.All']);
  const r = await fetch('https://graph.microsoft.com/v1.0/sites/provexpress.sharepoint.com:/sites/ProvexpressIntranet/comercial', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const d = await r.json();
  if(!d.id) throw new Error('Site no encontrado: ' + JSON.stringify(d.error||d));
  _siteId = d.id;
  console.log('[SITE OK]', _siteId);
  // Get drive ID for "Documentos compartidos"
  const r2 = await fetch('https://graph.microsoft.com/v1.0/sites/' + _siteId + '/drives', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const d2 = await r2.json();
  console.log('[DRIVES]', JSON.stringify(d2).slice(0,500));
  const drive = (d2.value||[]).find(dv =>
    dv.name === 'Documentos compartidos' || dv.name === 'Documents' || dv.name === 'Documentos'
  );
  if(drive) { _driveId = drive.id; console.log('[DRIVE OK]', drive.name, _driveId); }
  else { console.warn('[DRIVE] not found, using default'); }
  return _siteId;
}

async function loadDirectorFolder(siteId, folderName, token) {
  const authToken = token || await getToken(['Files.Read.All']);
  updateLoadingStatus('Cargando: ' + folderName + '...');
  const forecastBasePath = await getForecastBasePath(siteId, authToken);
  const folderPath = joinGraphPath(forecastBasePath, folderName);
  const driveBase = buildGraphRootUrl(siteId, folderPath, 'children?$top=50');
  console.log('[FOLDER]', driveBase);
  const r = await fetch(driveBase, { headers: { Authorization: 'Bearer ' + authToken } });
  const d = await r.json();
  if(!d.value) { console.warn('Sin archivos en', folderName, d); return; }
  const dirName = normalizeDirectorName(folderName.replace(/^(Grupo|Gupo)\s+/i,'').trim());
  if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
  const files = (d.value || []).filter(item =>
    item &&
    item.name &&
    item.name.match(/\.xlsx?$/i) &&
    !item.name.startsWith('~$') &&
    !item.name.toLowerCase().includes('base de datos')
  );
  await runWithConcurrencyLimit(files, FILE_LOAD_CONCURRENCY, async item => {
    updateLoadingStatus('Leyendo: ' + item.name);
    const bundle = await loadSpFileBundle(item, dirName);
    const recs = bundle.records || [];
    if(isSalesSupportFile(item.name)) {
      const meta = parseSalesSupportFileName(item.name) || {};
      const supportName = canonicalizeSalesSupportName(cleanNameSegment(meta.supportName || 'Sales Support'));
      const pendingRecords = ensureSalesPendingSupportName(bundle.pendingRecords || [], supportName);
      if(!LOADED_SALES_BY_SUPPORT[supportName]) LOADED_SALES_BY_SUPPORT[supportName] = [];
      LOADED_SALES_BY_SUPPORT[supportName].push({ name: item.name, dir: dirName, soporta: cleanNameSegment(meta.soportaName) });
      SALES_DATA.push(...recs);
      SALES_PENDING_DATA.push(...pendingRecords);
    } else {
      ALL_DATA.push(...recs);
      LOADED_FILES_BY_DIR[dirName].push(buildLoadedExcelFileMeta(item, dirName));
    }
  });
}

async function loadPreventaSpFile(item, dirName) {
  const url = item['@microsoft.graph.downloadUrl'];
  if(!url) return [];
  try {
    const buf = await (await fetch(url)).arrayBuffer();
    const wb = XLSX.read(buf, { type:'array', cellDates:true });
    const records = parseWorkbookMainRecords(wb, item, dirName, 'preventa');
    records.forEach(record => {
      record['PREVENTA'] = cleanDisplayText(record['COMERCIAL'], item.name.replace(/\.(xlsx|xls)$/i,'').trim());
    });
    console.log('[PREVENTA PARSE SP]', item.name, 'rows:', records.length);
    return records;
  } catch(e) {
    console.warn('Error leyendo preventa', item.name, e);
    return [];
  }
}

async function loadPreventaFolder(siteId, token) {
  const authToken = token || await getToken(['Files.Read.All']);
  const folderName = await getPreventaFolderName(siteId, authToken);
  updateLoadingStatus('Cargando: ' + folderName + '...');
  const forecastBasePath = await getForecastBasePath(siteId, authToken);
  const folderPath = joinGraphPath(forecastBasePath, folderName);
  const url = buildGraphRootUrl(siteId, folderPath, 'children?$top=50');
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + authToken } });
  const d = await r.json();
  if(!d.value) {
    console.warn('Sin archivos preventa en', folderName, d);
    return;
  }

  const files = (d.value || []).filter(item =>
    item &&
    item.name &&
    item.name.match(/\.xlsx?$/i) &&
    !item.name.startsWith('~$') &&
    !item.name.toLowerCase().includes('base de datos')
  );

  await runWithConcurrencyLimit(files, FILE_LOAD_CONCURRENCY, async item => {
    updateLoadingStatus('Leyendo preventa: ' + item.name);
    const recs = await loadPreventaSpFile(item, normalizeDirectorName(PREVENTA_FOLDER_NAME));
    const name = cleanNameSegment(item.name.replace(/\.(xlsx|xls)$/i,'').trim());
    LOADED_PREVENTA_FILES.push({ name, file: item.name });
    PREVENTA_DATA.push(...recs);
  });
}

function getExecTargetName() {
  const email = (CURRENT_USER && CURRENT_USER.email || '').toLowerCase().trim();
  const map = window.EXECUTIVO_BY_EMAIL || {};
  return (map[email] || CURRENT_USER.name || '').trim();
}

function buildExecSearchQueries(targetName, email) {
  const q = [];
  if(targetName) {
    q.push(targetName);
    const parts = targetName.split(' ').filter(Boolean);
    if(parts.length >= 2) q.push(parts[0] + ' ' + parts[1]);
  }
  if(email && email.includes('@')) {
    const local = email.split('@')[0].replace(/\./g,' ').trim();
    if(local) q.push(local);
  }
  return [...new Set(q.map(s => s.trim()).filter(s => s.length >= 3))];
}

function findBestExecFile(items, targetName) {
  const targetNorm = normalizePersonName(targetName||'');
  const cand = (items||[]).filter(it => it && it.name && /\.xlsx?$/i.test(it.name) && !it.name.startsWith('~$') && !isSalesSupportFile(it.name));
  if(!cand.length) return null;
  let file = cand.find(f => normalizePersonName(f.name) === targetNorm);
  if(!file) file = cand.find(f => namesMatch(f.name, targetName));
  if(!file && targetNorm) file = cand.find(f => normalizePersonName(f.name).includes(targetNorm));
  return file || null;
}

function buildSalesSupportSearchQueries(targetName, targetNames) {
  const names = [...new Set([targetName, ...(targetNames || [])].map(cleanNameSegment).filter(Boolean))];
  const queries = [];
  names.forEach(name => {
    queries.push('Sales Support ' + name);
    queries.push(name);
    const parts = name.split(' ').filter(Boolean);
    if(parts.length >= 3) queries.push('Sales Support ' + parts.slice(0, 3).join(' '));
    if(parts.length >= 2) {
      queries.push('Sales Support ' + parts.slice(0, 2).join(' '));
      queries.push(parts.slice(0, 2).join(' '));
    }
  });
  queries.push('Sales Support');
  return [...new Set(queries.map(s => s.trim()).filter(s => s.length >= 3))];
}

function findSalesSupportFiles(items, targetNames) {
  const candidates = (items || []).filter(item => {
    if(!item || !item.name || !/\.xlsx?$/i.test(item.name) || item.name.startsWith('~$')) return false;
    if(!isSalesSupportFile(item.name)) return false;
    const meta = parseSalesSupportFileName(item.name);
    return meta && matchesAnySalesSupportName(meta.supportName, targetNames);
  });

  return candidates.sort((a, b) => {
    const aMeta = parseSalesSupportFileName(a.name) || {};
    const bMeta = parseSalesSupportFileName(b.name) || {};
    const aName = normalizePersonName(aMeta.supportName || a.name);
    const bName = normalizePersonName(bMeta.supportName || b.name);
    return aName.localeCompare(bName) || String(a.name).localeCompare(String(b.name));
  });
}

function getDriveItemKey(item) {
  const driveId = item && item.parentReference && item.parentReference.driveId || item && item.__SEARCH_DRIVE_ID || '';
  return [driveId, item && item.id || '', item && item.webUrl || '', item && item.name || ''].join('|');
}

async function ensureDriveItemDownloadUrl(item, token) {
  if(!item || item['@microsoft.graph.downloadUrl']) return item;
  const driveId = item.parentReference && item.parentReference.driveId || item.__SEARCH_DRIVE_ID || _driveId;
  if(!driveId || !item.id) return item;
  const r = await fetch(buildDriveItemUrl(driveId, item.id), {
    headers: { Authorization: 'Bearer ' + token }
  });
  const d = await r.json().catch(() => ({}));
  if(!r.ok) {
    console.warn('[SALES SUPPORT SEARCH] no se pudo obtener downloadUrl', item.name, d);
    return item;
  }
  return {
    ...item,
    ...d,
    parentReference: {
      ...(item.parentReference || {}),
      ...(d.parentReference || {}),
      driveId
    },
    __SEARCH_DRIVE_ID: driveId,
    __SEARCH_DRIVE_NAME: item.__SEARCH_DRIVE_NAME
  };
}

async function loadSalesSupportSearchItem(item, token, targetName) {
  const file = await ensureDriveItemDownloadUrl(item, token);
  if(!file || !file['@microsoft.graph.downloadUrl']) return false;

  const meta = parseSalesSupportFileName(file.name) || {};
  const parentPath = file.parentReference && file.parentReference.path || '';
  const dirName = normalizeDirectorName(directorFromPath(parentPath)) || cleanDisplayText(file.__SEARCH_DRIVE_NAME, '');
  updateLoadingStatus('Leyendo: ' + file.name);
  const bundle = await loadSpFileBundle(file, dirName);
  const recs = bundle.records || [];
  const supportName = canonicalizeSalesSupportName(meta.supportName || targetName);
  const pendingRecords = ensureSalesPendingSupportName(bundle.pendingRecords || [], supportName);
  if(!LOADED_SALES_BY_SUPPORT[supportName]) LOADED_SALES_BY_SUPPORT[supportName] = [];
  LOADED_SALES_BY_SUPPORT[supportName].push({ name: file.name, dir: dirName, soporta: cleanNameSegment(meta.soportaName) });
  SALES_DATA.push(...recs);
  SALES_PENDING_DATA.push(...pendingRecords);
  return true;
}

async function searchDriveItemsGlobally(query, token) {
  const body = {
    requests: [{
      entityTypes: ['driveItem'],
      query: { queryString: query },
      from: 0,
      size: 50,
      fields: ['id','name','webUrl','parentReference','lastModifiedDateTime','lastModifiedBy','size']
    }]
  };
  const r = await fetch('https://graph.microsoft.com/v1.0/search/query', {
    method: 'POST',
    headers: graphJsonHeaders(token),
    body: JSON.stringify(body)
  });
  const d = await r.json().catch(() => ({}));
  if(!r.ok) {
    console.warn('[SALES SUPPORT SEARCH] busqueda global fallo', query, d);
    return [];
  }
  return (d.value || [])
    .flatMap(block => block.hitsContainers || [])
    .flatMap(container => container.hits || [])
    .map(hit => hit.resource)
    .filter(Boolean);
}

async function loadSalesSupportFilesBySearch(siteId, token, targetName, targetNames) {
  const queries = buildSalesSupportSearchQueries(targetName, targetNames);
  const drives = await getSearchDriveCandidates(siteId, token);
  const foundByKey = new Map();

  for(const drive of drives) {
    for(const query of queries) {
      try {
        const r = await fetch(buildDriveRootSearchUrl(drive.driveId, query), {
          headers: { Authorization: 'Bearer ' + token }
        });
        const d = await r.json().catch(() => ({}));
        if(!r.ok) {
          console.warn('[SALES SUPPORT SEARCH] fallo', drive.driveName, query, d);
          continue;
        }
        (d.value || []).forEach(item => {
          const enriched = {
            ...item,
            parentReference: {
              ...(item.parentReference || {}),
              driveId: item.parentReference && item.parentReference.driveId || drive.driveId
            },
            __SEARCH_DRIVE_ID: drive.driveId,
            __SEARCH_DRIVE_NAME: drive.driveName
          };
          foundByKey.set(getDriveItemKey(enriched), enriched);
        });
      } catch(e) {
        console.warn('[SALES SUPPORT SEARCH] error', drive.driveName, query, e);
      }
    }
  }

  for(const query of queries) {
    try {
      (await searchDriveItemsGlobally(query, token)).forEach(item => {
        foundByKey.set(getDriveItemKey(item), item);
      });
    } catch(e) {
      console.warn('[SALES SUPPORT SEARCH] error busqueda global', query, e);
    }
  }

  const files = findSalesSupportFiles(Array.from(foundByKey.values()), targetNames);
  if(!files.length) return false;

  let loaded = false;
  const loadedKeys = new Set();
  for(const item of files) {
    const key = getDriveItemKey(item);
    if(loadedKeys.has(key)) continue;
    loadedKeys.add(key);
    loaded = await loadSalesSupportSearchItem(item, token, targetName) || loaded;
  }
  return loaded;
}

async function loadSalesSupportFiles(siteId, token) {
  const filesToken = token || await getToken(['Files.Read.All']);
  const targetName = getSalesSupportTargetName();
  const targetNames = getSalesSupportTargetNames();
  let found = false;

  try {
    const forecastBasePath = await getForecastBasePath(siteId, filesToken);
    const folders = await getForecastFolders(siteId, filesToken);
    for(const folder of folders) {
      const folderPath = joinGraphPath(forecastBasePath, folder);
      try {
        const url = buildGraphRootUrl(siteId, folderPath, 'children?$top=100');
        const r = await fetch(url, { headers: { Authorization: 'Bearer ' + filesToken } });
        const d = await r.json();
        if(!d.value) continue;
        const dirName = normalizeDirectorName(folder.replace(/^(Grupo|Gupo)\s+/i,'').trim());
        for(const item of d.value) {
          if(!item.name.match(/\.xlsx?$/i)) continue;
          if(!isSalesSupportFile(item.name)) continue;
          const meta = parseSalesSupportFileName(item.name);
          if(!meta || !matchesAnySalesSupportName(meta.supportName, targetNames)) continue;
          updateLoadingStatus('Leyendo: ' + item.name);
          const bundle = await loadSpFileBundle(item, dirName);
          const recs = bundle.records || [];
          const supportName = canonicalizeSalesSupportName(meta.supportName || targetName);
          const pendingRecords = ensureSalesPendingSupportName(bundle.pendingRecords || [], supportName);
          if(!LOADED_SALES_BY_SUPPORT[supportName]) LOADED_SALES_BY_SUPPORT[supportName] = [];
          LOADED_SALES_BY_SUPPORT[supportName].push({ name: item.name, dir: dirName, soporta: cleanNameSegment(meta.soportaName) });
          SALES_DATA.push(...recs);
          SALES_PENDING_DATA.push(...pendingRecords);
          found = true;
        }
      } catch(e) { console.warn('Error leyendo folder sales', folder, e); }
    }
  } catch(e) {
    console.warn('[SALES SUPPORT] no se pudo listar FORECAST 2026, se usara busqueda por archivo', e);
  }

  if(!found) {
    updateLoadingStatus('Buscando archivo Sales Support...');
    found = await loadSalesSupportFilesBySearch(siteId, filesToken, targetName, targetNames);
  }

  if(!found) {
    throw new Error('No se encontraron archivos Sales Support para ' + (targetName || 'este usuario') + ' por carpeta ni por busqueda.');
  }
}

async function searchExecFileInForecast(siteId, targetName, token) {
  const email = (CURRENT_USER && CURRENT_USER.email || '').toLowerCase().trim();
  const queries = buildExecSearchQueries(targetName, email);
  if(!queries.length) return null;
  const authToken = token || await getToken(['Files.Read.All']);
  const folderPath = await getForecastBasePath(siteId, authToken);
  for(const q of queries) {
    const url = buildGraphRootUrl(siteId, folderPath, 'search(q=\'' + encodeURIComponent(q) + '\')?$top=50');
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + authToken } });
    const d = await r.json();
    const file = findBestExecFile(d.value || [], targetName);
    if(file) return file;
  }
  return null;
}

async function loadEjecutivoFile(siteId, token) {
  const filesToken = token || await getToken(['Files.Read.All']);
  const forecastBasePath = await getForecastBasePath(siteId, filesToken);
  const folders = await getForecastFolders(siteId, filesToken);
  const targetName = getExecTargetName();
  let found = false;
  for(const folder of folders) {
    try {
      const folderPath = joinGraphPath(forecastBasePath, folder);
      const r = await fetch(buildGraphRootUrl(siteId, folderPath, 'children?$top=50'), { headers: { Authorization: 'Bearer ' + filesToken } });
      const d = await r.json();
      if(!d.value) continue;
      const file = findBestExecFile(d.value, targetName);
      if(file) {
        const readyFile = await ensureDriveItemDownloadUrl(file, filesToken);
        const dirName = normalizeDirectorName(folder.replace(/^(Grupo|Gupo)\s+/i,'').trim());
        if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
        const recs = await loadSpFile(readyFile, dirName);
        ALL_DATA.push(...recs);
        LOADED_FILES_BY_DIR[dirName].push(buildLoadedExcelFileMeta(readyFile, dirName));
        found = true;
        return true;
      }
    } catch(e) { continue; }
  }
  // Fallback: search in FORECAST 2026 if file is nested or renamed
  const fallback = await searchExecFileInForecast(siteId, targetName, filesToken);
  if(fallback) {
    const readyFallback = await ensureDriveItemDownloadUrl(fallback, filesToken);
    const dirName = normalizeDirectorName(directorFromPath((fallback.parentReference && fallback.parentReference.path) || '') || '');
    if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
    const recs = await loadSpFile(readyFallback, dirName);
    ALL_DATA.push(...recs);
    LOADED_FILES_BY_DIR[dirName].push(buildLoadedExcelFileMeta(readyFallback, dirName));
    return true;
  }
  if(!found) {
    throw new Error('No se encontró el Excel de ' + (targetName || 'este usuario') + '. Verifica el nombre del archivo en FORECAST 2026.');
  }
}

function parseWorkbookMainRecords(wb, item, dirName, datasetType){
  const wsName = pickWorksheetName(wb.SheetNames, datasetType);
  const ws = wb.Sheets[wsName];
  if(!ws) return [];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  const hdrIdx = findHeaderRowIndex(raw);
  if(hdrIdx<0) return [];
  const hdrs = raw[hdrIdx].map(h=>h ? mapHeaderName(String(h).trim()) : '');
  const recs = [];
  for(let i=hdrIdx+1;i<raw.length;i++) {
    const row=raw[i]; if(!isMeaningfulDataRow(row)) continue;
    const rec={};
    hdrs.forEach((h,j)=>{ if(h) rec[h]=row[j]!==undefined?row[j]:null; });
    if(!hasMeaningfulRecordContent(rec)) continue;
    decorateRecordFromFile(rec, item.name, dirName);
    registerRecord(rec, item.name, wsName, datasetType);
    recs.push(rec);
  }
  if(datasetType === 'sales') console.log('[SALES PARSE SP]', item.name, wsName, 'rows:', recs.length, 'headerRow:', hdrIdx);
  return recs;
}

function parseWorkbookSalesPendingRecords(wb, item, dirName){
  if(!isSalesSupportFile(item.name)) return [];
  const wsName = pickSalesPendingWorksheetName(wb.SheetNames);
  if(!wsName) {
    console.warn('[SALES PENDIENTES] hoja no encontrada', item.name, wb.SheetNames);
    return [];
  }
  const ws = wb.Sheets[wsName];
  if(!ws) return [];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  const hdrIdx = findSalesPendingHeaderRowIndex(raw);
  if(hdrIdx<0) {
    const preview = raw.slice(0, 8).map(row =>
      (row || []).map(cell => normalizeHeaderKey(mapHeaderName(cell))).filter(Boolean)
    );
    console.warn('[SALES PENDIENTES] encabezado no encontrado', item.name, wsName, preview);
    return [];
  }
  const hdrs = raw[hdrIdx].map(h=>h ? mapHeaderName(String(h).trim()) : '');
  const recs = [];
  for(let i=hdrIdx+1;i<raw.length;i++) {
    const row=raw[i]; if(!isMeaningfulDataRow(row)) continue;
    const rec={};
    hdrs.forEach((h,j)=>{ if(h) rec[h]=row[j]!==undefined?row[j]:null; });
    if(!hasMeaningfulSalesPendingRecordContent(rec)) continue;
    decorateSalesPendingRecordFromFile(rec, item.name, dirName);
    registerRecord(rec, item.name, wsName, 'sales_pending');
    recs.push(rec);
  }
  console.log('[SALES PENDIENTES PARSE SP]', item.name, wsName, 'rows:', recs.length, 'headerRow:', hdrIdx);
  return recs;
}

async function loadSpFileBundle(item, dirName) {
  const url = item['@microsoft.graph.downloadUrl'];
  if(!url) return { records: [], pendingRecords: [] };
  try {
    const buf = await (await fetch(url)).arrayBuffer();
    const wb  = XLSX.read(buf, { type:'array', cellDates:true });
    const datasetType = isSalesSupportFile(item.name) ? 'sales' : 'forecast';
    return {
      records: parseWorkbookMainRecords(wb, item, dirName, datasetType),
      pendingRecords: datasetType === 'sales' ? parseWorkbookSalesPendingRecords(wb, item, dirName) : []
    };
  } catch(e) { console.warn('Error leyendo', item.name, e); return { records: [], pendingRecords: [] }; }
}

async function loadSpFile(item, dirName) {
  const bundle = await loadSpFileBundle(item, dirName);
  return bundle.records;
}

// ── Tabs por rol ─────────────────────────────
function applyRoleTabs() {
  if(!CURRENT_USER) return;
  const { role } = CURRENT_USER;
  const tabs = {
    gerencia:  document.getElementById('tab-gerencia'),
    director:  document.getElementById('tab-director'),
    ejecutivo: document.getElementById('tab-ejecutivo'),
    sales:     document.getElementById('tab-sales'),
    preventa:  document.getElementById('tab-preventa'),
    divisas:   document.getElementById('tab-divisas'),
    marcas:    document.getElementById('tab-marcas'),
    resumen:   document.getElementById('tab-resumen'),
    finanzas:  document.getElementById('tab-finanzas'),
  };
  // Reset — mostrar todas
  Object.values(tabs).forEach(t => { if(t) t.style.display = ''; });

  if(role === 'sales_support') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.director && (tabs.director.style.display = 'none');
    tabs.ejecutivo && (tabs.ejecutivo.style.display = 'none');
    tabs.preventa && (tabs.preventa.style.display = 'none');
    tabs.divisas  && (tabs.divisas.style.display  = 'none');
    tabs.marcas   && (tabs.marcas.style.display   = 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
    tabs.finanzas && (tabs.finanzas.style.display = 'none');
    showPage('sales', tabs.sales);
  } else if(role === 'ejecutivo') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.director && (tabs.director.style.display = 'none');
    tabs.sales    && (tabs.sales.style.display    = 'none');
    tabs.preventa && (tabs.preventa.style.display = 'none');
    tabs.divisas  && (tabs.divisas.style.display  = 'none');
    tabs.marcas   && (tabs.marcas.style.display   = 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
    tabs.finanzas && (tabs.finanzas.style.display = 'none');
    showPage('ejecutivo', tabs.ejecutivo);
  } else if(role === 'director') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.ejecutivo&& (tabs.ejecutivo.style.display= 'none');
    tabs.preventa && (tabs.preventa.style.display = 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
    tabs.finanzas && (tabs.finanzas.style.display = 'none');
    showPage('director', tabs.director);
  } else {
    // gerencia / gerencia_director — ven todo
    showPage('gerencia', tabs.gerencia);
  }
}

// ── Auto-cargar al abrir desde SharePoint ────
window.addEventListener('DOMContentLoaded', () => {
  // Cargar TRM automáticamente
  fetchTRM();
  // Pre-cargar MSAL y auto-login siempre
  loadMSAL(() => {
    initMsalApp();
    // Auto-cargar siempre al abrir
    loadFolderFromSharePoint();
  });
});

// Exponer handlers usados por atributos inline (onclick/onchange) en index.html
window.loadFolderFromSharePoint = loadFolderFromSharePoint;
window.showPage = showPage;
window.toggleAppTheme = toggleAppTheme;
window.renderDivisas = renderDivisas;
window.renderFinanzas = renderFinanzas;
window.setFinancePeriod = setFinancePeriod;
window.setFinanceEndDate = setFinanceEndDate;
window.saveFinanceGoal = saveFinanceGoal;
window.loadFinanceSummary = loadFinanceSummary;
window.setDivisaEstadoFilter = setDivisaEstadoFilter;
window.showMoreDivisaRows = showMoreDivisaRows;
window.showMoreMarcasBars = showMoreMarcasBars;
window.renderDirector = renderDirector;
window.renderEjecutivo = renderEjecutivo;
window.renderSales = renderSales;
window.renderPreventa = renderPreventa;
window.setSalesView = setSalesView;
window.setSalesPendingFilter = setSalesPendingFilter;
window.renderMarcas = renderMarcas;
window.downloadMarcasLineasExcel = downloadMarcasLineasExcel;
window.downloadExecutiveExcelMetadataReport = downloadExecutiveExcelMetadataReport;
window.selectEjecutivo = selectEjecutivo;
window.clearEjecutivoBrandFocus = clearEjecutivoBrandFocus;
window.openExecNegociosFromMarcas = openExecNegociosFromMarcas;
window.selectSalesSupport = selectSalesSupport;
window.selectPreventa = selectPreventa;
window.openMarcaLineaDetail = openMarcaLineaDetail;
window.closeMarcaLineaDetail = closeMarcaLineaDetail;
window.setMarcaLineaDetailEstado = setMarcaLineaDetailEstado;
window.setDirectorEstadoFilter = setDirectorEstadoFilter;
window.showMoreEstadoRows = showMoreEstadoRows;
window.openNegocioDetailById = openNegocioDetailById;
window.closeNegocioDetail = closeNegocioDetail;
