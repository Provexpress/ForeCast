/* ══════════════════════════════════════
   DATA & STATE
══════════════════════════════════════ */
let ALL_DATA = [];
let SALES_DATA = [];
let TRM = 4150;
let TRM_READY = false;
let SELECTED_EXEC_BY_DIR = {};
let RECORD_SEQ = 0;
let NEGOCIO_DETAIL_STATE = null;
let MARCA_LINEA_DETAIL_STATE = null;
let LOADED_SALES_BY_SUPPORT = {};

const GERENCIA_ESTADO_STEP = 20;
const GERENCIA_ESTADO_LIMITS = {
  GANADA: 30,
  PENDIENTE: 30,
  PERDIDA: 30,
  APLAZADO: 30
};
const DIVISAS_DETAIL_STEP = 10;
const DIVISAS_DETAIL_LIMITS = { COP: 10, USD: 10 };
const TOP_BAR_LIMIT = 15;

const TRM_CACHE_KEY = 'trm_last';
const TRM_DATE_KEY = 'trm_date';

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
function cacheTRM(val, dateStr){
  try {
    if(val && val > 100){
      localStorage.setItem(TRM_CACHE_KEY, String(val));
      if(dateStr) localStorage.setItem(TRM_DATE_KEY, String(dateStr));
    }
  } catch(_) {}
}
async function fetchText(url, timeoutMs){
  const ctrl = new AbortController();
  const t = setTimeout(()=>ctrl.abort(), timeoutMs || 8000);
  try {
    const r = await fetch(url, {cache:'no-store', signal: ctrl.signal});
    if(!r.ok) throw new Error('HTTP '+r.status);
    return await r.text();
  } finally { clearTimeout(t); }
}
async function fetchFirstText(urls, timeoutMs){
  for(const u of urls){
    try { return await fetchText(u, timeoutMs); } catch(_) {}
  }
  return '';
}
function extractTRMFromText(text){
  const reHtml = new RegExp('<span class=\"valor-indicador-principal\">\\s*([0-9.]+,[0-9]{2})\\s*COP\\/USD<\\/span>[\\s\\S]*?<span class=\"fecha-indicador-principal\">\\s*(\\d{2}\\/\\d{2}\\/\\d{4})<\\/span>');
  const m = text.match(reHtml);
  if(m){
    return {value: parseFloat(m[1].replace(/\\./g,'').replace(',','.')), date: m[2]};
  }
  const rePlain = new RegExp('([0-9.]+,[0-9]{2})\\s*COP\\/USD');
  const mp = text.match(rePlain);
  if(mp){
    const dm = text.match(new RegExp('(\\d{2}\\/\\d{2}\\/\\d{4})'));
    return {value: parseFloat(mp[1].replace(/\\./g,'').replace(',','.')), date: dm ? dm[1] : ''};
  }
  return null;
}
function extractTRMFromSDMX(text){
  const mv = text.match(/ObsValue[^>]*value=\"([0-9.]+)\"/);
  if(!mv) return null;
  const md = text.match(/ObsDimension[^>]*value=\"(\\d{8})\"/);
  let dateStr = '';
  if(md && md[1] && md[1].length===8){
    dateStr = md[1].substring(6,8)+'/'+md[1].substring(4,6)+'/'+md[1].substring(0,4);
  }
  return {value: parseFloat(mv[1]), date: dateStr};
}

loadCachedTRM();

async function fetchTRM() {
  const setTRM = t => {
    TRM = t;
    TRM_READY = true;
    const inp = document.getElementById('trm-input');
    if(inp) inp.value = Number(TRM).toFixed(2);
    if(ALL_DATA.length || SALES_DATA.length) renderAll();
    cacheTRM(TRM, window._trmDate);
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
function fmtTRM(v){
  if(v===null||v===undefined||isNaN(v)) return '—';
  return Number(v).toLocaleString('es-CO',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function fmtTRMDisplay(v){
  return TRM_READY ? fmtTRM(v) : '—';
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
  return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
}

function jsStringLiteral(s){
  return JSON.stringify(String(s ?? ''));
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

function normalizeHeaderKey(v){
  return String(v || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toUpperCase()
    .replace(/\s+/g,' ')
    .trim();
}

const HEADER_KEY_MAP = {
  'VENTA CLIENTE':'MONTO VENTA CLIENTE',
  'MONTO VENTA':'MONTO VENTA CLIENTE',
  'VALOR VENTA':'MONTO VENTA CLIENTE',
  'VALOR VENTA CLIENTE':'MONTO VENTA CLIENTE',
  'VALOR CLIENTE':'MONTO VENTA CLIENTE',
  'NUMERO COTIZACION':'NUMERO DE COTIZACION',
  'NRO COTIZACION':'NUMERO DE COTIZACION',
  'NO COTIZACION':'NUMERO DE COTIZACION',
  'COTIZACION':'NUMERO DE COTIZACION',
  'NUMERO PARTE':'NUMERO DE PARTE',
  'NRO PARTE':'NUMERO DE PARTE',
  'NO PARTE':'NUMERO DE PARTE',
  'PART NUMBER':'NUMERO DE PARTE',
  'COSTO DEL NEGOCIO':'COSTO NEGOCIO',
  'COSTO TOTAL NEGOCIO':'COSTO NEGOCIO',
  'MONEDA':'MONEDA 2',
  'DIVISA':'MONEDA 2',
  'MONEDA2':'MONEDA 2',
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
  'NOTAS':'OBSERVACIONES'
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
  OBSERVACIONES: 'Observaciones'
};

function mapHeaderName(h){
  const raw = String(h || '').trim();
  const normalized = normalizeHeaderKey(raw);
  return HEADER_KEY_MAP[normalized] || raw;
}

function registerRecord(rec, sourceFile, sourceSheet, dataset){
  const ds = dataset === 'sales' ? 'sales' : 'forecast';
  rec.__RID = (ds === 'sales' ? 'sales-' : 'neg-') + (++RECORD_SEQ);
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

function getRowBrandName(row){
  return firstFilled(row, ['MARCA','FABRICANTE']) || '';
}

function getRowPartNumber(row){
  return firstFilled(row, ['NUMERO DE PARTE','NUMERO PARTE','NRO PARTE','NO PARTE','PART NUMBER']) || '';
}

function getRowLineName(row){
  return firstFilled(row, ['LINEA DE PRODUCTO','LINEA']) || '';
}

function getRowObservation(row){
  return firstFilled(row, ['OBSERVACIONES','OBSERVACION','COMENTARIOS','COMENTARIO']);
}

function getSalesSupportName(row){
  return firstFilled(row, ['SALES SUPPORT','COMERCIAL']) || '';
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
  return ALL_DATA.find(r => r.__RID === id) || SALES_DATA.find(r => r.__RID === id) || null;
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

function isSalesSupportFile(name){
  const base = String(name || '').replace(/\.(xlsx|xls)$/i,'').trim();
  return /^sales support\b/i.test(base);
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

function findHeaderRowIndex(rows){
  for(let i = 0; i < rows.length; i++) {
    if(isLikelyHeaderRow(rows[i])) return i;
  }
  return -1;
}

function parseSalesSupportFileName(name){
  const base = String(name || '').replace(/\.(xlsx|xls)$/i,'').trim();
  if(!/^sales support\b/i.test(base)) return null;
  const rest = base.replace(/^sales support\b/i,'').trim();
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
  rec['DIRECTOR'] = directorHint ? toTitleName(directorHint) : toTitleName(rec['DIRECTOR'] || rec['DIRECTOR '] || '');
  rec['ESTADO'] = normalizeEstado(rec['ESTADO']);
  if(salesMeta){
    const supportName = cleanNameSegment(salesMeta.supportName);
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
  if(!sel) return;
  const current = sel.value;
  sel.innerHTML = `<option value="">Todos</option>${months.map(month => `<option value="${month}">${getMonthLongLabel(month)}</option>`).join('')}`;
  if(current && months.includes(current)) sel.value = current;
  else sel.value = '';
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
/* ══════════════════════════════════════
   FOLDER & FILE LOADING — detecta estructura director/ejecutivo
══════════════════════════════════════ */

// Parsea un archivo Excel y devuelve array de registros
function parseXlsx(file, directorHint){
  return new Promise((resolve)=>{
    const r = new FileReader();
    r.onload = function(ev){
      try{
        const wb = XLSX.read(ev.target.result,{type:'binary',cellDates:true});
        const datasetType = isSalesSupportFile(file.name) ? 'sales' : 'forecast';
        const wsName = pickWorksheetName(wb.SheetNames, datasetType);
        const ws = wb.Sheets[wsName];
        const raw = XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
        
        const hdrIdx = findHeaderRowIndex(raw);
        if(hdrIdx<0){resolve([]);return;}
        
        const rawHdrs = raw[hdrIdx].map(h=>h?String(h).trim():'');
        const hdrs = rawHdrs.map(mapHeaderName);
        const recs = [];
        // Debug: log headers for first file
        if(!window._hdrDebugDone){ window._hdrDebugDone=true; console.log('[HDRS]', file.name, hdrs.filter(h=>h)); }
        for(let i=hdrIdx+1;i<raw.length;i++){
          const row=raw[i];
          if(!row[1]) continue;
          const rec={};
          hdrs.forEach((h,j)=>{if(h) rec[h]=row[j]!==undefined?row[j]:null;});
          decorateRecordFromFile(rec, file.name, directorHint);
          registerRecord(rec, file.name, wsName, datasetType);
          if(recs.length===0) console.log('[ROW1]', file.name, {fecha:rec['FECHA DIA/MES/AÑO'], monto:rec['MONTO VENTA CLIENTE'], moneda:rec['MONEDA 2'], cliente:rec['CLIENTE']});
          recs.push(rec);
        }
        resolve(recs);
      }catch(err){console.warn('Error parseando',file.name,err);resolve([]);}
    };
    r.readAsBinaryString(file);
  });
}

// Extrae el nombre del director desde el path de la carpeta
// Ej: "FORECAST 2026/Grupo Juan David Novoa/Freddy.xlsx" → "Juan David Novoa"
function directorFromPath(path){
  const parts = path.split('/');
  // Buscar la parte que empiece con "Grupo" o "Gupo" (typo en sharepoint)
  const dirPart = parts.find(p=> /^(Grupo|Gupo)/i.test(p.trim()));
  if(dirPart){
    return dirPart.replace(/^(Grupo|Gupo)\s+/i,'').trim();
  }
  // Si no, usar el folder padre del archivo
  if(parts.length>=2) return parts[parts.length-2].trim();
  return '';
}

// Procesa lista de File objects (de folder o archivos sueltos)
let LOADED_FILES_BY_DIR = {}; // track all loaded files even if empty

async function processFiles(files){
  // Filtrar solo .xlsx/.xls
  const xlsFiles = files.filter(f=>{
    if(!/\.(xlsx|xls)$/i.test(f.name)) return false;
    if(f.name.startsWith('~$')) return false; // Excel temp files
    // Ignorar archivos que claramente no son ejecutivos (bases de datos, plantillas, etc.)
    const nameLower = f.name.toLowerCase();
    if(nameLower.includes('base de datos')) return false;
    if(nameLower.includes('plantilla')) return false;
    if(nameLower.includes('template')) return false;
    if(/^\d/.test(f.name)) return false; // empieza con número
    return true;
  });
  if(!xlsFiles.length){alert('No se encontraron archivos .xlsx de ejecutivos en la carpeta seleccionada');return;}
  
  ALL_DATA = [];
  SALES_DATA = [];
  RECORD_SEQ = 0;
  LOADED_SALES_BY_SUPPORT = {};
  
  // Mostrar overlay de progreso (funciona tanto si upload-zone está visible como oculto)
  let overlay = document.getElementById('load-overlay');
  if(!overlay){
    overlay = document.createElement('div');
    overlay.id = 'load-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(3,5,14,.88);backdrop-filter:blur(8px);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px';
    overlay.innerHTML = `
      <div style="font-family:var(--font-display);font-size:13px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--corp-cyan)">Cargando datos...</div>
      <div id="load-status" style="font-size:12px;color:var(--text3);font-family:var(--font-body);min-width:320px;text-align:center"></div>
      <div style="background:var(--border);border-radius:6px;height:6px;overflow:hidden;width:320px">
        <div id="load-bar" style="height:100%;background:linear-gradient(90deg,var(--corp-blue),var(--corp-cyan));width:0%;transition:width .25s;border-radius:6px"></div>
      </div>
      <div id="folder-structure" style="max-width:480px;text-align:left"></div>
    `;
    document.body.appendChild(overlay);
  } else {
    overlay.style.display = 'flex';
  }
  const progBar  = document.getElementById('load-bar');
  const progTxt  = document.getElementById('load-status');
  const fstruc   = document.getElementById('folder-structure');
  fstruc.innerHTML = '';
  
  // Agrupar por director según path
  const byDir = {};
  xlsFiles.forEach(f=>{
    const path = f.webkitRelativePath || f.name;
    const dir  = directorFromPath(path) || 'Sin Director';
    if(!byDir[dir]) byDir[dir]=[];
    byDir[dir].push(f);
  });
  LOADED_FILES_BY_DIR = {};
  
  // Procesar uno a uno con progreso
  let done=0;
  const total=xlsFiles.length;
  const dirColors={'0':'#2D7BF7','1':'#00C2D4','2':'#0DBF82','3':'#F0A020','4':'#E040A0'};
  
  const dirNames = Object.keys(byDir);
  for(let di=0;di<dirNames.length;di++){
    const dirName = dirNames[di];
    for(const f of byDir[dirName]){
      progTxt.textContent = `Leyendo: ${f.name} (${done+1}/${total})`;
      progBar.style.width = Math.round(((done)/total)*100)+'%';
      const recs = await parseXlsx(f, dirName);
      if(isSalesSupportFile(f.name)){
        const meta = parseSalesSupportFileName(f.name) || {};
        const supportName = cleanNameSegment(meta.supportName || 'Sales Support');
        if(!LOADED_SALES_BY_SUPPORT[supportName]) LOADED_SALES_BY_SUPPORT[supportName] = [];
        LOADED_SALES_BY_SUPPORT[supportName].push({ name: f.name, dir: dirName, soporta: cleanNameSegment(meta.soportaName) });
        SALES_DATA.push(...recs);
      } else {
        if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
        LOADED_FILES_BY_DIR[dirName].push(f);
        ALL_DATA.push(...recs);
      }
      done++;
    }
  }
  progBar.style.width = '100%';
  progTxt.textContent = `✅ ${total} archivo${total!==1?'s':''} cargados — ${ALL_DATA.length} forecast / ${SALES_DATA.length} sales`;
  
  // Mostrar estructura detectada
  let structHTML = '<div style="font-size:10px;font-family:var(--font-display);font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--text2);margin-bottom:10px">Estructura detectada:</div>';
  dirNames.forEach((d,di)=>{
    const c = Object.values(dirColors)[di%5];
    const regularFiles = byDir[d].filter(f => !isSalesSupportFile(f.name));
    const salesFiles = byDir[d].filter(f => isSalesSupportFile(f.name));
    const ejecs = [...new Set(regularFiles.map(f=>f.name.replace(/\.(xlsx|xls)$/i,'')))];
    const sales = [...new Set(salesFiles.map(f=>cleanNameSegment((parseSalesSupportFileName(f.name)||{}).supportName || f.name.replace(/\.(xlsx|xls)$/i,''))))];
    structHTML += `<div style="margin-bottom:8px">
      <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:${c};font-family:var(--font-display);font-weight:700">
        <span>📁</span> ${d} <span style="color:var(--text2);font-weight:400">(${ejecs.length} ejecutivo${ejecs.length!==1?'s':''}${sales.length ? ' · '+sales.length+' sales' : ''})</span>
      </div>
      <div style="margin-left:20px;margin-top:4px;display:flex;flex-wrap:wrap;gap:6px">
        ${ejecs.map(e=>`<span style="font-size:10px;background:${c}15;border:1px solid ${c}30;border-radius:4px;padding:2px 8px;color:var(--text3)">📄 ${e}</span>`).join('')}
        ${sales.map(e=>`<span style="font-size:10px;background:rgba(42,191,223,.12);border:1px solid rgba(42,191,223,.28);border-radius:4px;padding:2px 8px;color:#CFEFFF">🤝 ${e}</span>`).join('')}
      </div>
    </div>`;
  });
  fstruc.innerHTML = structHTML;
  fstruc.style.display = 'block';
  
  // Ocultar overlay y mostrar dashboard
  setTimeout(()=>{
    const ov = document.getElementById('load-overlay');
    if(ov) ov.style.display='none';
    // Asegurarse que el dashboard esté visible
    const uz = document.getElementById('upload-zone-g');
    if(uz) uz.style.display='none';
    document.getElementById('gerencia-content').style.display='block';
    document.getElementById('hoy-strip').style.display='block';
    finalizeLoad();
  }, 1200);
}

// TRM change
document.getElementById('trm-input').addEventListener('input',function(){
  TRM_READY = true;
  TRM=getTRM();
  if(ALL_DATA.length || SALES_DATA.length) renderAll();
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
    const nDirs = Object.keys(LOADED_FILES_BY_DIR).length;
    reloadInfo.textContent = CURRENT_USER && CURRENT_USER.role === 'sales_support'
      ? nSalesFiles + ' archivos sales · Última carga: ' + new Date().toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'})
      : nFiles + ' archivos cargados · ' + nDirs + ' equipos' + (nSalesFiles ? ' · '+nSalesFiles+' sales' : '') + ' · Última carga: ' + new Date().toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'});
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
  const dirs = [...new Set(visibleData.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))].sort();
  const execs = [...new Set(visibleData.map(r=>r['COMERCIAL']||'').filter(Boolean))].sort();
  const execsWithData = [...new Set(visibleData.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  const salesSupports = [...new Set(visibleSales.map(r=>getSalesSupportName(r)).filter(Boolean))];
  const salesTargets = [...new Set(visibleSales.map(r=>getSalesSoportaName(r)).filter(Boolean))];
  document.getElementById('file-count-hd').textContent =
    CURRENT_USER && CURRENT_USER.role === 'sales_support'
      ? `${visibleSales.length} registros sales · ${salesSupports.length} support · ${salesTargets.length} apoyos`
      : `${visibleData.length} negocios · ${dirs.length} dir · ${execsWithData.length} ejecutivos${visibleSales.length ? ' · '+visibleSales.length+' sales' : ''}`;
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
  // Aplicar pestañas y badge según rol
  applyRoleTabs();
  showUserBadge();

  // Populate selects
  const selDir=document.getElementById('sel-director');
  const dirsForSel=[...new Set([
    ...visibleData.map(r=>(r['DIRECTOR']||'').trim()),
    ...ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()),
    ...Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim())
  ].filter(Boolean))].sort();
  selDir.innerHTML=dirsForSel.map(d=>`<option value="${d}">${d}</option>`).join('');
  
  const execsFromFiles2=Object.values(LOADED_FILES_BY_DIR||{}).flat()
    .map(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim()).filter(Boolean);
  const allExecsForSel=[...new Set([...execs,...execsFromFiles2])].sort();
  const selEj=document.getElementById('sel-ejecutivo');
  selEj.innerHTML=allExecsForSel.map(e=>`<option value="${e}">${e}</option>`).join('');

  renderAll();
}

function getVisibleData() {
  if(!CURRENT_USER) return ALL_DATA;
  const { role, directorGroup, name } = CURRENT_USER;
  if(role === 'sales_support') return [];
  if(role === 'director') {
    return ALL_DATA.filter(r => (r['DIRECTOR']||'').trim() === (directorGroup||'').trim());
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

function getSalesSupportTargetName() {
  const email = (CURRENT_USER && CURRENT_USER.email || '').toLowerCase().trim();
  const map = window.SALES_SUPPORT_BY_EMAIL || {};
  return (map[email] || CURRENT_USER && CURRENT_USER.name || '').trim();
}

function getVisibleSalesData() {
  if(!CURRENT_USER) return SALES_DATA;
  const { role, directorGroup } = CURRENT_USER;
  if(role === 'sales_support') {
    const targetName = getSalesSupportTargetName();
    return SALES_DATA.filter(r => namesMatch(getSalesSupportName(r), targetName));
  }
  if(role === 'director') {
    return SALES_DATA.filter(r => (r['DIRECTOR']||'').trim() === (directorGroup||'').trim());
  }
  if(role === 'ejecutivo') return [];
  return SALES_DATA;
}

function renderAll(){
  refreshForecastMonthFilters();
  renderGerencia();
  renderDirector();
  renderEjecutivo();
  renderSales();
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
  renderGerenciaEstadoTables(getVisibleData());
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
  return getVisibleData().filter(row => normalizeCategoryValue(getMarcaLineaDetailRowValue(row, state.type)) === target);
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
  renderMarcaLineaDetail();
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
  renderNegocioDetail(row);
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
  if(mapped === 'MONTO VENTA CLIENTE' || mapped === 'COSTO NEGOCIO' || mapped === 'COSTO' || mapped === 'UTILIDAD'){
    const mon = cleanDisplayText(row['MONEDA 2'], 'COP').toUpperCase();
    const monto = parseMonto(value);
    return mon === 'USD' ? fmtUSD(monto) : fmtCOP(monto);
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

      <div class="detail-kpi-grid">
        <div class="detail-stat">
          <span>Total negocios</span>
          <strong>${fmtNum(totalNegocios)}</strong>
          <small>Relacionados con ${escHtml(state.value)}</small>
        </div>
        <div class="detail-stat">
          <span>Ganadas</span>
          <strong>${fmtNum(estados[0].count)}</strong>
          <small>${fmtCOP(estados[0].total)}</small>
        </div>
        <div class="detail-stat">
          <span>Pendientes</span>
          <strong>${fmtNum(estados[1].count)}</strong>
          <small>${fmtCOP(estados[1].total)}</small>
        </div>
        <div class="detail-stat">
          <span>Perdidas</span>
          <strong>${fmtNum(estados[2].count)}</strong>
          <small>${fmtCOP(estados[2].total)}</small>
        </div>
        <div class="detail-stat">
          <span>Aplazadas</span>
          <strong>${fmtNum(estados[3].count)}</strong>
          <small>${fmtCOP(estados[3].total)}</small>
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
            <div class="chart-hd">Negocios asociados</div>
            <div style="font-size:11px;color:var(--text2)">Filtrados por ${escHtml(typeLabel.toLowerCase())} y estado del negocio.</div>
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
          ${buildTable(filteredRows, { clickable:true, sourcePage:'marca-linea-detail' })}
        </div>
      </div>
    </div>
  `;
}

/* ══════════════════════════════════════
   CHART HELPERS
══════════════════════════════════════ */
function renderBars(containerId, items, color, fmtFn, opts){
  const el=document.getElementById(containerId);
  if(!el) return;
  const options = opts || {};
  const max=Math.max(...items.map(i=>i.val),1);
  el.innerHTML=items.map((it,idx)=>{
    const pct=Math.round((it.val/max)*100);
    const c=Array.isArray(color)?color[idx%color.length]:color;
    const onClick = options.getOnClick ? options.getOnClick(it, idx) : '';
    const rowAttrs = onClick
      ? ` class="bar-row bar-row-action" role="button" tabindex="0" onclick="${escAttr(onClick)}" onkeydown="${escAttr(`if(event.key==='Enter'||event.key===' '){event.preventDefault();${onClick}}`)}" title="${escAttr(options.clickTitle || 'Abrir detalle')}" data-tooltip="${escAttr(options.tooltipPrefix ? options.tooltipPrefix + it.name : 'Abrir detalle de ' + it.name)}"`
      : ` class="bar-row"`;
    return `<div${rowAttrs}>
      <div class="bar-name w100" title="${it.name}">${it.name}</div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${pct}%;background:${c}20;border:1px solid ${c}40">
          <span class="bar-val" style="color:${c}">${fmtFn?fmtFn(it.val):abr(it.val)}</span>
        </div>
      </div>
    </div>`;
  }).join('');
}

function renderDonut(svgId, legId, items){
  const svg=document.getElementById(svgId);
  const leg=document.getElementById(legId);
  if(!svg||!leg) return;
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
    paths+=`<path d="M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} Z" fill="${c}" opacity=".85"/>`;
    angle+=slice;
  });
  svg.innerHTML=paths+`<circle cx="50" cy="50" r="22" fill="#0B0F1E"/>`;
  
  leg.innerHTML=items.slice(0,6).map((it,i)=>{
    const pct=((it.val/total)*100).toFixed(1);
    return `<div class="leg-item" style="display:flex;align-items:center;gap:8px;font-size:12px"><div class="leg-dot" style="background:${COLORS[i%COLORS.length]};width:10px;height:10px;border-radius:50%;flex-shrink:0"></div><span title="${it.name}" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text-strong);font-weight:600">${it.name}</span><span class="leg-pct" style="color:var(--text);font-family:var(--font-mono);font-weight:600">${pct}%</span></div>`;
  }).join('');
}

function renderEvoChart(containerId, dataByDir, months){
  const el=document.getElementById(containerId);
  if(!el) return;
  const monthKeys = (months && months.length)
    ? months
    : [...new Set(Object.values(dataByDir || {}).flatMap(row => Object.keys(row || {})).filter(Boolean))].sort();
  if(!monthKeys.length){
    el.innerHTML = `<div style="font-size:11px;color:var(--text2)">Sin datos mensuales para mostrar.</div>`;
    return;
  }
  const dirs=[...new Set(Object.keys(dataByDir))];
  const nDirs=dirs.length;
  const allVals=dirs.flatMap(d=>monthKeys.map(m=>dataByDir[d][m]||0));
  const maxVal=Math.max(...allVals,1);
  const W=520,H=200,padL=56,padR=12,padT=20,padB=36;
  const gW=W-padL-padR, gH=H-padT-padB;
  const grpW=gW/monthKeys.length;
  const barW=Math.min(18, (grpW-8)/Math.max(nDirs,1));
  const gap=2;

  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;overflow:visible">`;
  svg+=`<defs>${dirs.map((d,di)=>`<linearGradient id="bg${di}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${COLORS[di%COLORS.length]}" stop-opacity=".95"/><stop offset="100%" stop-color="${COLORS[di%COLORS.length]}" stop-opacity=".4"/></linearGradient>`).join('')}</defs>`;

  // Grid lines
  [0,.25,.5,.75,1].forEach(t=>{
    const y=padT+gH*(1-t);
    svg+=`<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="#1A2240" stroke-width="${t===0?1.5:.7}"/>`;
    if(t>0) svg+=`<text x="${padL-5}" y="${y+3.5}" text-anchor="end" font-size="8.5" font-weight="400" fill="#B8C8E8" font-family="IBM Plex Mono,monospace">${abr(maxVal*t)}</text>`;
  });

  // Bars
  monthKeys.forEach((m,mi)=>{
    const grpCenter=padL+(mi+0.5)*grpW;
    const totalBarW=(barW+gap)*nDirs-gap;
    dirs.forEach((d,di)=>{
      const v=dataByDir[d][m]||0;
      const bh=Math.max(v/maxVal*gH, v>0?2:0);
      const x=grpCenter - totalBarW/2 + di*(barW+gap);
      const y=padT+gH-bh;
      svg+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW}" height="${bh.toFixed(1)}" rx="2" fill="url(#bg${di})" data-tooltip="${escAttr(dirs[di]+': '+abr(v))}"></rect>`;
    });
    // Month label
    svg+=`<text x="${grpCenter.toFixed(1)}" y="${H-8}" text-anchor="middle" font-size="10" fill="#B8C8E8" font-family="IBM Plex Sans,sans-serif" font-weight="400">${getMonthShortLabel(m)}</text>`;
  });

  // Legend top
  let lx=padL;
  dirs.forEach((d,di)=>{
    const c=COLORS[di%COLORS.length];
    const short=d.split(' ').slice(0,2).join(' ');
    svg+=`<rect x="${lx}" y="4" width="9" height="9" rx="2" fill="${c}"/>`;
    svg+=`<text x="${lx+12}" y="12" font-size="9" font-weight="400" fill="#B8C8E8" font-family="IBM Plex Sans,sans-serif">${short}</text>`;
    lx+=short.length*4.8+20;
  });

  svg+='</svg>';
  el.innerHTML=svg;
  attachChartTooltips(el);
}

/* ══════════════════════════════════════
   GERENCIA GENERAL
══════════════════════════════════════ */
function renderGerencia(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
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
      <div class="kpi-val">${fmtUSD(totalUSD)}</div>
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
  renderBars('bar-directores',dirData,COLORS);
  
  // Evo por director mensual
  const monthsForEvo = getForecastMonths(ALL_DATA);
  const dirsForEvo=[...new Set([
    ...ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()),
    ...Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim())
  ].filter(Boolean))].sort();
  const evoByDir={};
  dirsForEvo.forEach(d=>{
    evoByDir[d]={};
    monthsForEvo.forEach(m=>{
      evoByDir[d][m]=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d&&getMonth(r['FECHA DIA/MES/AÑO'])===m).reduce((s,r)=>s+toCOP(r),0);
    });
  });
  renderEvoChart('evo-dir-chart',evoByDir, monthsForEvo);
  
  // Bar ejecutivos
  const execs=[...new Set(ALL_DATA.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  const ejData=execs.map(e=>({name:e.split(' ')[0],val:ALL_DATA.filter(r=>r['COMERCIAL']===e).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderBars('bar-ejecutivos',ejData.slice(0, TOP_BAR_LIMIT),COLORS);
  
  // Donuts
  const estados=['GANADA','PENDIENTE','PERDIDA','APLAZADO'];
  const estData=estados.map(e=>({name:e,val:ALL_DATA.filter(r=>r['ESTADO']===e).reduce((s,r)=>s+toCOP(r),0)}));
  renderDonut('donut-estado','leg-estado',estData);
  
  const lineas=[...new Set(ALL_DATA.map(r=>r['LINEA DE PRODUCTO']||'').filter(Boolean))];
  const linData=lineas.map(l=>({name:l,val:ALL_DATA.filter(r=>r['LINEA DE PRODUCTO']===l).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderDonut('donut-linea','leg-linea',linData);

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
        <tr class="table-row-action" style="border-top:1px solid var(--border)" onclick="openNegocioDetailById('${r.__RID}','gerencia')" title="Abrir detalle del negocio">
          <td style="padding:5px 8px;font-size:10px;color:var(--text);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escAttr(cliente)}">${escHtml(cliente)}</td>
          <td style="padding:5px 8px;font-size:10px;color:#B0BCDF;font-weight:600;white-space:nowrap">${escHtml(comercial)}</td>
          <td style="padding:5px 8px;font-size:10px;color:#B0BCDF;font-weight:500;white-space:nowrap;max-width:80px;overflow:hidden;text-overflow:ellipsis" title="${escAttr(linea)}">${escHtml(linea)}</td>
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
            <span style="font-size:9px;color:#B0BCDF;margin-left:8px;font-family:var(--font-body)">${rows.length} negocio${rows.length!==1?'s':''}</span>
          </div>
          <span style="font-family:var(--font-mono);font-size:13px;font-weight:700;color:var(--text)">${abr(total)}</span>
        </div>
        <div class="estado-card-body" style="overflow-y:auto;max-height:280px" data-scroll-key="estado:${estado}">
          <table class="estado-mini-table" style="width:100%;border-collapse:collapse">
            <thead style="position:sticky;top:0;z-index:1;background:var(--bg2)">
              <tr>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:#C8D4F0;text-align:left">EMPRESA</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:#C8D4F0;text-align:left">EJECUTIVO</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:#C8D4F0;text-align:left">LÍNEA</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:#C8D4F0;text-align:right">VALOR</th>
              </tr>
            </thead>
            <tbody>${rows_html}</tbody>
            <tfoot style="background:var(--bg2);border-top:1px solid var(--border)">
              <tr>
                <td colspan="3" style="padding:6px 8px;font-size:9px;font-family:var(--font-display);color:#C8D4F0;font-weight:700">TOTAL ${rows.length} NEGOCIOS</td>
                <td style="padding:6px 8px;font-size:11px;text-align:right;font-family:var(--font-mono);color:${colores[estado]};font-weight:700">${abr(total)}</td>
              </tr>
            </tfoot>
          </table>
        </div>${more}
      </div>`;
  }).join('');
}

/* ══════════════════════════════════════
   DIRECTOR
══════════════════════════════════════ */
function renderDirector(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const dir=document.getElementById('sel-director').value;
  const mes=document.getElementById('sel-dir-mes').value;
  const est=document.getElementById('sel-dir-estado').value;
  const trm=getTRM();
  const estadoOptions = [
    { value:'', label:'Todos' },
    { value:'GANADA', label:'GANADA' },
    { value:'PENDIENTE', label:'PENDIENTE' },
    { value:'PERDIDA', label:'PERDIDA' },
    { value:'APLAZADO', label:'APLAZADO' }
  ].map(opt => `<option value="${opt.value}" ${opt.value===est?'selected':''}>${opt.label}</option>`).join('');
  const detailToolbar = `
    <div class="director-table-toolbar">
      <div>
        <div class="director-table-toolbar-label">Filtro rapido de tabla</div>
        <div class="director-table-toolbar-meta">Estado sincronizado con el filtro superior del director</div>
      </div>
      <div class="filter-group director-table-filter">
        <div class="filter-label">Estado tabla</div>
        <select id="sel-dir-estado-detail" onchange="setDirectorEstadoFilter(this.value)">
          ${estadoOptions}
        </select>
      </div>
    </div>`;

  let data=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir);
  if(mes) data=data.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])===mes);
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
  data.forEach(r=>{const m=getMonth(r['FECHA DIA/MES/AÑO']);if(evoMonths[m]!==undefined)evoMonths[m]+=toCOP(r);});
  
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
      s+=`<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="#1A2240" stroke-width="${t===0?1.5:.6}"/>`;
      if(t>0) s+=`<text x="${padL-4}" y="${y+3}" text-anchor="end" font-size="7.5" fill="#B8C8E8" font-weight="400" font-family="IBM Plex Mono,monospace">${abr(maxV*t)}</text>`;
    });
    months.forEach((m,i)=>{
      const v=vals[i]||0;
      const bh=Math.max(v/maxV*gH, v>0?2:0);
      const cx=padL+(i+0.5)*(gW/months.length);
      const x=cx-bW/2;
      const y=padT+gH-bh;
      s+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bW}" height="${bh.toFixed(1)}" rx="3" fill="url(#bg-dir)" data-tooltip="${escAttr(getMonthShortLabel(m)+': '+abr(v))}"></rect>`;
      s+=`<text x="${cx.toFixed(1)}" y="${H-6}" text-anchor="middle" font-size="9" fill="#B0BCDF" font-family="IBM Plex Sans,sans-serif" font-weight="400">${getMonthShortLabel(m)}</text>`;
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
    return `<div class="kpi exec-card ${isSelected?'selected':''}" style="--ac:${c};min-width:160px;opacity:${hasD?1:.55}" onclick="selectDirectorExec('${e}')">
      <div class="kpi-accent"></div>
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
        <div style="width:30px;height:30px;border-radius:50%;background:${c}30;border:1px solid ${c}60;display:flex;align-items:center;justify-content:center;font-family:var(--font-display);font-size:11px;font-weight:700;color:${c}">${ini}</div>
        <div>
          <div style="font-size:11px;font-family:var(--font-display);font-weight:700;color:var(--text)">${e.split(' ')[0]}</div>
          <div style="font-size:9px;color:var(--text3)">${hasD?ejData.length+' negocios':'Sin datos aún'}</div>
        </div>
      </div>
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
      <h2>${selectedExec}</h2>
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
        <div class="kpi-val">${fmtUSD(execTotalUSD)}</div>
        <div class="kpi-sub">Liq: ${abr(execTotalUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad COP</div>
        <div class="kpi-val">${abr(execUtilidadCOP)}</div>
        <div class="kpi-sub">${fmtCOP(execUtilidadCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad USD</div>
        <div class="kpi-val">${fmtUSD(execUtilidadUSD)}</div>
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
        <div class="kpi-sub">Director: ${dir}</div>
      </div>
    </div>
    <div class="chart-card g1">
      <div class="chart-hd">Detalle de Negocios — ${selectedExec}</div>
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
    <div class="section-hd"><h2>${dir}</h2><span class="section-tag">DIRECTOR</span></div>
    
    <div class="kpi-grid kpi-grid-6" style="margin-bottom:16px">
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        <div class="kpi-val">${fmtUSD(totalUSD)}</div>
        <div class="kpi-sub">Liq: ${abr(totalUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad COP</div>
        <div class="kpi-val">${abr(utilidadCOP)}</div>
        <div class="kpi-sub">${fmtCOP(utilidadCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad USD</div>
        <div class="kpi-val">${fmtUSD(utilidadUSD)}</div>
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
        <div class="chart-hd">Evolución Mensual — ${dir.split(' ')[0]}</div>
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

    <div class="section-hd"><h2>Equipo</h2><span class="section-tag">${execs.length} EJECUTIVOS</span></div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:18px">
      ${ejCards}
    </div>
    ${execKPIs}
  `;
  attachChartTooltips(document.getElementById('director-content'));
  
  // Donut estado director — usar datos sin filtro de estado para mostrar distribución real
  const dataForDonut=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir && (!mes||getMonth(r['FECHA DIA/MES/AÑO'])===mes));
  const estD=['GANADA','PENDIENTE','PERDIDA','APLAZADO'].map(e=>({name:e,val:dataForDonut.filter(r=>r['ESTADO']===e).reduce((s,r)=>s+toCOP(r),0)}));
  renderDonut('donut-dir-est','leg-dir-est',estD);
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
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const role = CURRENT_USER ? CURRENT_USER.role : null;
  const targetName = role === 'ejecutivo' ? getExecTargetName() : '';
  const selEj = document.getElementById('sel-ejecutivo');
  const mes=document.getElementById('sel-ej-mes').value;
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
      selEj.innerHTML = `<option value="${targetName}">${targetName}</option>`;
      selEj.value = targetName;
    } else {
      const current = selEj.value;
      selEj.innerHTML = allExecs.map(e=>`<option value="${e}">${e}</option>`).join('');
      if(current && allExecs.includes(current)) selEj.value = current;
      else if(allExecs[0]) selEj.value = allExecs[0];
    }
  }
  const ej = (role === 'ejecutivo' && targetName) ? targetName : (selEj ? selEj.value : '');
  const execs = (role === 'ejecutivo' && targetName) ? [targetName] : allExecs;
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
    return `<div class="persona-card ${selected} ${hasData?'':'no-data'}" onclick="selectEjecutivo('${e}')">
      <div class="persona-avatar" style="background:${c}${hasData?'25':'10'};border:2px solid ${c}${hasData?'50':'20'};color:${hasData?c:'var(--text2)'}">${initials(e)}</div>
      <div class="persona-name" style="color:${hasData?'var(--text)':'var(--text2)'}">${e}</div>
      <div class="persona-role">${dir}</div>
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
  if(mes) data=data.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])===mes);
  if(est) data=data.filter(r=>r['ESTADO']===est);
  
  const totalCOP=data.reduce((s,r)=>s+toCOP(r),0);
  const totalUSD=data.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const ganadas=data.filter(r=>r['ESTADO']==='GANADA');
  const ejColor=COLORS[execs.indexOf(ej)%COLORS.length];
  
  const lineas=[...new Set(data.map(r=>r['LINEA DE PRODUCTO']||'').filter(Boolean))];
  const linData=lineas.map(l=>({name:l,val:data.filter(r=>r['LINEA DE PRODUCTO']===l).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  
  document.getElementById('ejecutivo-content').innerHTML=`
    <div class="section-hd" style="margin-top:16px"><h2>${ej}</h2><span class="section-tag" style="background:${ejColor}20;color:${ejColor};border-color:${ejColor}40">EJECUTIVO</span></div>
    
    <div class="kpi-grid kpi-grid-4" style="margin-bottom:16px">
      <div class="kpi" style="--ac:${ejColor}"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        <div class="kpi-val">${fmtUSD(totalUSD)}</div>
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
        <div class="kpi-sub">Director: ${(data[0]||{})['DIRECTOR']||'—'}</div>
      </div>
    </div>
    
    <div class="g2">
      <div class="chart-card">
        <div class="chart-hd">Top Líneas — ${ej.split(' ')[0]}</div>
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
      <div class="chart-hd">Detalle de Negocios — ${ej}</div>
      ${buildTable(data, { clickable:true, sourcePage:'ejecutivo' })}
    </div>
  `;
  
  renderBars('bar-ej-lineas',linData,COLORS);
  const estD=['GANADA','PENDIENTE','PERDIDA','APLAZADO'].map(e=>({name:e,val:data.filter(r=>r['ESTADO']===e).length}));
  renderDonut('donut-ej-est','leg-ej-est',estD);
}

function selectEjecutivo(name){
  document.getElementById('sel-ejecutivo').value=name;
  renderEjecutivo();
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
      const estadoClass = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'].includes(estado) ? estado : 'PENDIENTE';
      const rowAttrs = options.clickable && r.__RID
        ? ` class="table-row-action" onclick="openNegocioDetailById('${r.__RID}','${escAttr(sourcePage)}')" title="Abrir detalle del registro sales"`
        : '';
      return `<tr${rowAttrs}>
        <td class="td-mono" data-label="Fecha">${escHtml(fecha)}</td>
        <td class="td-mono" data-label="Cotizacion">${escHtml(cotizacion)}</td>
        <td data-label="Cliente">${escHtml(cliente)}</td>
        <td data-label="Soporta">${escHtml(soporta)}</td>
        <td data-label="Moneda"><span class="badge ${mon==='USD'?'badge-PEDIDA':'badge-PENDIENTE'}">${mon}</span></td>
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

function renderSales(){
  const allSales = getVisibleSalesData();
  const role = CURRENT_USER ? CURRENT_USER.role : null;
  const targetName = role === 'sales_support' ? getSalesSupportTargetName() : '';
  const host = document.getElementById('sales-content');
  const grid = document.getElementById('sales-support-grid');
  const selSupport = document.getElementById('sel-sales-support');
  const selMes = document.getElementById('sel-sales-mes');
  const selEstado = document.getElementById('sel-sales-estado');
  if(!host || !grid || !selSupport || !selMes || !selEstado) return;

  const supportsFromData = [...new Set(allSales.map(r=>getSalesSupportName(r)).filter(Boolean))];
  const supportsFromFiles = Object.keys(LOADED_SALES_BY_SUPPORT || {}).filter(Boolean);
  const allSupports = [...new Set([...supportsFromData, ...supportsFromFiles])].sort((a,b)=>a.localeCompare(b,'es'));

  if(role === 'sales_support' && targetName) {
    selSupport.innerHTML = `<option value="${escAttr(targetName)}">${escHtml(targetName)}</option>`;
    selSupport.value = targetName;
  } else {
    const current = selSupport.value;
    selSupport.innerHTML = allSupports.map(name=>`<option value="${escAttr(name)}">${escHtml(name)}</option>`).join('');
    if(current && allSupports.includes(current)) selSupport.value = current;
    else if(allSupports[0]) selSupport.value = allSupports[0];
  }

  const monthValues = [...new Set(allSales.map(r=>getMonth(getRowDateValue(r))).filter(Boolean))].sort();
  const currentMonth = selMes.value;
  selMes.innerHTML = `<option value="">Todos</option>${monthValues.map(month=>`<option value="${month}">${getMonthLabel(month)}</option>`).join('')}`;
  if(currentMonth && monthValues.includes(currentMonth)) selMes.value = currentMonth;

  if(!allSupports.length){
    grid.innerHTML = '';
    host.innerHTML = `<div class="chart-card g1"><div style="font-size:12px;color:var(--text2)">No hay archivos Sales Support cargados para este usuario o grupo.</div></div>`;
    return;
  }

  const selectedSupport = role === 'sales_support' && targetName ? targetName : selSupport.value;
  const mes = selMes.value;
  const estado = selEstado.value;

  grid.innerHTML = allSupports.map((supportName, idx)=>{
    const supportRows = allSales.filter(r => namesMatch(getSalesSupportName(r), supportName));
    const totalCOP = supportRows.reduce((sum,row)=>sum+toCOP(row),0);
    const totalGanadas = supportRows.filter(r=>cleanDisplayText(r['ESTADO'],'').toUpperCase()==='GANADA').length;
    const supportedCount = [...new Set(supportRows.map(r=>getSalesSoportaName(r)).filter(Boolean))].length;
    const c = COLORS[idx % COLORS.length];
    const hasData = supportRows.length > 0;
    const selected = namesMatch(selectedSupport, supportName) ? 'selected' : '';
    return `<div class="persona-card ${selected} ${hasData?'':'no-data'}" onclick="${escAttr(`selectSalesSupport(${jsStringLiteral(supportName)})`)}">
      <div class="persona-avatar" style="background:${c}${hasData?'25':'10'};border:2px solid ${c}${hasData?'50':'20'};color:${hasData?c:'var(--text2)'}">${initials(supportName)}</div>
      <div class="persona-name" style="color:${hasData?'var(--text)':'var(--text2)'}">${escHtml(supportName)}</div>
      <div class="persona-role">Sales Support</div>
      ${hasData
        ? `<div class="persona-stats">
            <div class="p-stat"><div class="p-stat-label">Total</div><div class="p-stat-val" style="color:${c};font-size:11px">${abr(totalCOP)}</div></div>
            <div class="p-stat"><div class="p-stat-label">Registros</div><div class="p-stat-val">${supportRows.length}</div></div>
            <div class="p-stat"><div class="p-stat-label">Ganadas</div><div class="p-stat-val" style="color:var(--corp-green)">${totalGanadas}</div></div>
            <div class="p-stat"><div class="p-stat-label">Soporta</div><div class="p-stat-val">${supportedCount}</div></div>
          </div>`
        : `<div style="font-size:9px;color:var(--text3);font-family:var(--font-display);margin-top:8px;padding:5px 8px;background:rgba(255,255,255,.03);border-radius:6px;letter-spacing:.5px">Sin registros aun</div>`}
    </div>`;
  }).join('');

  let data = allSales.filter(r => namesMatch(getSalesSupportName(r), selectedSupport));
  if(mes) data = data.filter(r => getMonth(getRowDateValue(r)) === mes);
  if(estado) data = data.filter(r => cleanDisplayText(r['ESTADO'],'').toUpperCase() === estado);
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
    <div class="kpi-grid kpi-grid-6" style="margin-bottom:16px">
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        <div class="kpi-val">${fmtUSD(totalUSD)}</div>
        <div class="kpi-sub">TRM dia: ${fmtTRMDisplay(getTRM())}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-purple2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad COP</div>
        <div class="kpi-val">${abr(utilidadCOP)}</div>
        <div class="kpi-sub">${fmtCOP(utilidadCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
        <div class="kpi-label">Utilidad USD</div>
        <div class="kpi-val">${fmtUSD(utilidadUSD)}</div>
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

  renderBars('bar-sales-soporta', topSoporta, COLORS);
  renderDonut('donut-sales-est', 'leg-sales-est', estadoData);
}

function selectSalesSupport(name){
  const sel = document.getElementById('sel-sales-support');
  if(sel) sel.value = name;
  renderSales();
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
  
  const usdData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='USD');
  const copData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='COP');
  const usdDetailData=(estadoDetalle ? usdData.filter(r=>r['ESTADO']===estadoDetalle) : usdData).sort((a,b)=>toCOP(b)-toCOP(a));
  const copDetailData=(estadoDetalle ? copData.filter(r=>r['ESTADO']===estadoDetalle) : copData).sort((a,b)=>toCOP(b)-toCOP(a));
  const usdVisible = DIVISAS_DETAIL_LIMITS.USD || 10;
  const copVisible = DIVISAS_DETAIL_LIMITS.COP || 10;
  const usdRemaining = Math.max(usdDetailData.length - usdVisible, 0);
  const copRemaining = Math.max(copDetailData.length - copVisible, 0);
  
  const totalUSD=usdData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const totalCOP=copData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const usdLiqCOP=totalUSD*trm;
  const granTotal=totalCOP+usdLiqCOP;
  
  document.getElementById('divisas-cards').innerHTML=`
    <div class="divisa-card cop">
      <div class="divisa-label" style="color:var(--cop-color)"><span class="flag">🇨🇴</span> COP — Peso Colombiano</div>
      <div class="divisa-main" style="color:var(--cop-color)">${abr(totalCOP)}</div>
      <div class="divisa-sub">${fmtCOP(totalCOP)}</div>
      <div class="divisa-stats">
        <div><div class="d-stat-label">Negocios</div><div class="d-stat-val" style="color:var(--cop-color)">${copData.length}</div></div>
        <div><div class="d-stat-label">Ganadas</div><div class="d-stat-val" style="color:var(--corp-green)">${copData.filter(r=>r['ESTADO']==='GANADA').length}</div></div>
        <div><div class="d-stat-label">Prom. Negocio</div><div class="d-stat-val" style="color:var(--cop-color)">${abr(totalCOP/(copData.length||1))}</div></div>
      </div>
    </div>
    <div class="divisa-card usd">
      <div class="divisa-label" style="color:var(--usd-color)"><span class="flag">🇺🇸</span> USD — Dólar Americano</div>
      <div class="divisa-main" style="color:var(--usd-color)">${fmtUSD(totalUSD)}</div>
      <div class="divisa-sub">TRM ${fmtTRMDisplay(trm)} → Liquidado ${abr(usdLiqCOP)}</div>
      <div class="divisa-stats">
        <div><div class="d-stat-label">Negocios</div><div class="d-stat-val" style="color:var(--usd-color)">${usdData.length}</div></div>
        <div><div class="d-stat-label">Ganadas</div><div class="d-stat-val" style="color:var(--corp-green)">${usdData.filter(r=>r['ESTADO']==='GANADA').length}</div></div>
        <div><div class="d-stat-label">En COP</div><div class="d-stat-val" style="color:var(--usd-color)">${abr(usdLiqCOP)}</div></div>
      </div>
    </div>
  `;
  
  // Table USD detail
  document.getElementById('tbl-usd').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Cliente</th><th>Producto</th><th>USD</th><th>COP Liquidado</th><th>Estado</th></tr></thead>
    <tbody>${usdDetailData.length ? usdDetailData.slice(0, usdVisible).map(r=>{
      const usd=parseMonto(r['MONTO VENTA CLIENTE'])||0;
      const liq=usd*trm;
      return `<tr>
        <td>${(r['COMERCIAL']||'').split(' ')[0]}</td>
        <td>${r['CLIENTE']||'—'}</td>
        <td style="max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r['PRODUCTO']||'—'}</td>
        <td class="td-mono td-usd">${fmtUSD(usd)}</td>
        <td class="td-mono td-cop">${fmtCOP(liq)}</td>
        <td><span class="badge badge-${r['ESTADO']}">${r['ESTADO']||'—'}</span></td>
      </tr>`;
    }).join('') : `<tr><td colspan="6" style="text-align:center;color:var(--text2);padding:20px 14px">Sin negocios ${estadoDetalle ? estadoDetalle.toLowerCase() : ''} en USD.</td></tr>`}</tbody>
  </table>${usdRemaining > 0 ? `<div class="table-more-wrap"><button type="button" class="table-more-btn" onclick="showMoreDivisaRows('USD')">Ver mas (${usdRemaining})</button></div>` : ''}`;
  
  // Table COP detail
  document.getElementById('tbl-cop').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Cliente</th><th>Producto</th><th>COP</th><th>Estado</th></tr></thead>
    <tbody>${copDetailData.length ? copDetailData.slice(0, copVisible).map(r=>{
      const cop=parseMonto(r['MONTO VENTA CLIENTE'])||0;
      return `<tr>
        <td>${(r['COMERCIAL']||'').split(' ')[0]}</td>
        <td>${r['CLIENTE']||'—'}</td>
        <td style="max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r['PRODUCTO']||'—'}</td>
        <td class="td-mono td-cop">${fmtCOP(cop)}</td>
        <td><span class="badge badge-${r['ESTADO']}">${r['ESTADO']||'—'}</span></td>
      </tr>`;
    }).join('') : `<tr><td colspan="5" style="text-align:center;color:var(--text2);padding:20px 14px">Sin negocios ${estadoDetalle ? estadoDetalle.toLowerCase() : ''} en COP.</td></tr>`}</tbody>
  </table>${copRemaining > 0 ? `<div class="table-more-wrap"><button type="button" class="table-more-btn" onclick="showMoreDivisaRows('COP')">Ver mas (${copRemaining})</button></div>` : ''}`;
  
  // Tabla resumen consolidado
  const dirs=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))];
  document.getElementById('tbl-resumen-divisas').innerHTML=`<table>
    <thead><tr><th>Director</th><th>Negos COP</th><th>Valor COP</th><th>Negos USD</th><th>Valor USD</th><th>Liq. USD→COP</th><th>TOTAL COP</th></tr></thead>
    <tbody>${dirs.map(d=>{
      const dd=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d);
      const dc=dd.filter(r=>(r['MONEDA 2']||'').trim()==='COP');
      const du=dd.filter(r=>(r['MONEDA 2']||'').trim()==='USD');
      const vCOP=dc.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const vUSD=du.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const liq=vUSD*trm;
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:700;color:var(--text)">${d}</td>
        <td class="td-mono">${dc.length}</td>
        <td class="td-mono td-cop">${fmtCOP(vCOP)}</td>
        <td class="td-mono">${du.length}</td>
        <td class="td-mono td-usd">${fmtUSD(vUSD)}</td>
        <td class="td-mono td-usd">${fmtCOP(liq)}</td>
        <td class="td-mono" style="color:#fff;font-weight:600">${fmtCOP(vCOP+liq)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td style="font-family:var(--font-display);font-weight:800;color:#fff">TOTAL GENERAL</td>
      <td class="td-mono">${copData.length}</td>
      <td class="td-mono td-cop" style="font-weight:700">${fmtCOP(totalCOP)}</td>
      <td class="td-mono">${usdData.length}</td>
      <td class="td-mono td-usd" style="font-weight:700">${fmtUSD(totalUSD)}</td>
      <td class="td-mono td-usd" style="font-weight:700">${fmtCOP(usdLiqCOP)}</td>
      <td class="td-mono" style="color:var(--corp-cyan);font-weight:800;font-size:13px">${fmtCOP(granTotal)}</td>
    </tr></tbody>
  </table>`;
}

/* ══════════════════════════════════════
   MARCAS
══════════════════════════════════════ */
function renderMarcas(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  
  const marcas=[...new Set(ALL_DATA.map(r=>r['MARCA']||'').filter(Boolean))];
  const marcaData=marcas.map(m=>({name:m,val:ALL_DATA.filter(r=>r['MARCA']===m).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderBars('bar-marcas',marcaData.slice(0, TOP_BAR_LIMIT),COLORS,null,{
    getOnClick: item => `openMarcaLineaDetail('marca', ${jsStringLiteral(item.name)}, 'marcas')`,
    clickTitle: 'Abrir detalle de marca'
  });
  renderDonut('donut-marca','leg-marca',marcaData);
  
  const lineas=[...new Set(ALL_DATA.map(r=>r['LINEA DE PRODUCTO']||'').filter(Boolean))];
  const linData=lineas.map(l=>({name:l,val:ALL_DATA.filter(r=>r['LINEA DE PRODUCTO']===l).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderBars('bar-lineas',linData,COLORS,null,{
    getOnClick: item => `openMarcaLineaDetail('linea', ${jsStringLiteral(item.name)}, 'marcas')`,
    clickTitle: 'Abrir detalle de linea'
  });
  renderDonut('donut-linea2','leg-linea2',linData);
  
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
      ? [`<option value="">Todos los directores</option>`, ...directorOptions.map(d=>`<option value="${escAttr(d)}">${escHtml(d)}</option>`)]
      : directorOptions.map(d=>`<option value="${escAttr(d)}">${escHtml(d)}</option>`);
    directorSelect.innerHTML = opts.join('');
    if(prev && directorOptions.includes(prev)) directorSelect.value = prev;
    else if(directorOptions.length === 1) directorSelect.value = directorOptions[0];
    else if(directorOptions.length > 1) directorSelect.value = directorOptions[0];
    selectedDirector = directorSelect.value || '';
  }
  if(directorFilterWrap) directorFilterWrap.style.display = directorOptions.length > 1 ? '' : 'none';

  const marcaExecData = selectedDirector
    ? ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===selectedDirector)
    : ALL_DATA;
  const execs=[...new Set(marcaExecData.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  document.getElementById('tbl-marca-ej').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th>${marcas.map(m=>`<th>${m}</th>`).join('')}<th>Top Marca</th></tr></thead>
    <tbody>${execs.length ? execs.map(e=>{
      const ed=marcaExecData.filter(r=>r['COMERCIAL']===e);
      const marcaCounts=marcas.map(m=>ed.filter(r=>r['MARCA']===m).length);
      const topIdx=marcaCounts.indexOf(Math.max(...marcaCounts));
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)">${e.split(' ')[0]}</td>
        ${marcaCounts.map((c,i)=>`<td class="td-mono" style="color:${c>0?COLORS[i%COLORS.length]:'var(--text2)'}">${c||'—'}</td>`).join('')}
        <td style="font-family:var(--font-display);font-weight:700;color:${COLORS[topIdx%COLORS.length]}">${marcas[topIdx]||'—'}</td>
      </tr>`;
    }).join('') : `<tr><td colspan="${marcas.length+2}" style="text-align:center;color:var(--text2);padding:20px 14px">Sin ejecutivos con datos para este grupo.</td></tr>`}</tbody>
  </table>`;
  
  // Productos únicos
  const prods=[...new Set(ALL_DATA.map(r=>r['PRODUCTO']||'').filter(Boolean))];
  document.getElementById('tbl-productos').innerHTML=`<table>
    <thead><tr><th>Producto / Tipo Máquina</th><th>Marca</th><th>Línea</th><th>Moneda</th><th>Negocios</th><th>Valor Total COP</th></tr></thead>
    <tbody>${prods.map(p=>{
      const pd=ALL_DATA.filter(r=>r['PRODUCTO']===p);
      const tot=pd.reduce((s,r)=>s+toCOP(r),0);
      const marca=pd[0]?pd[0]['MARCA']||'—':'—';
      const linea=pd[0]?pd[0]['LINEA DE PRODUCTO']||'—':'—';
      const moneda=(r=>r['MONEDA 2']||'—')(pd[0]||{});
      return `<tr class="table-row-action" onclick="${escAttr(`openMarcaLineaDetail('producto', ${jsStringLiteral(p)}, 'marcas')`)}" title="Abrir detalle del producto">
        <td style="color:var(--text)">${p}</td>
        <td class="td-mono" style="color:var(--corp-cyan)">${marca}</td>
        <td>${linea}</td>
        <td><span class="badge ${moneda==='USD'?'badge-PEDIDA':'badge-PENDIENTE'}">${moneda}</span></td>
        <td class="td-mono">${pd.length}</td>
        <td class="td-mono td-cop">${fmtCOP(tot)}</td>
      </tr>`;
    }).join('')}</tbody>
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
  
  document.getElementById('tbl-total-cop').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Director</th><th>Negocios COP</th>${monthHeaderCells}<th>Total COP</th></tr></thead>
    <tbody>${execs.map(e=>{
      const ed=copData.filter(r=>r['COMERCIAL']===e);
      const dir=ALL_DATA.find(r=>r['COMERCIAL']===e);
      const monthlyValues = months.map(month => ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])===month).reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0));
      const tot=monthlyValues.reduce((sum, val)=>sum+val,0);
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)">${e}</td>
        <td style="color:var(--text2)">${dir?dir['DIRECTOR']||'—':'—'}</td>
        <td class="td-mono">${ed.length}</td>
        ${monthlyValues.map(val=>`<td class="td-mono td-cop">${val>0?abr(val):'—'}</td>`).join('')}
        <td class="td-mono td-cop" style="font-weight:700">${fmtCOP(tot)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;color:var(--cop-color)">SUBTOTAL COP</td>
      <td class="td-mono">${copData.length}</td>
      ${months.map(month => `<td class="td-mono td-cop">${abr(copData.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])===month).reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0))}</td>`).join('')}
      <td class="td-mono td-cop" style="font-weight:800;font-size:13px">${fmtCOP(totalCOP)}</td>
    </tr></tbody>
  </table>`;
  
  // Tabla USD detail + liquidado
  document.getElementById('tbl-total-usd').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Director</th><th>Negocios USD</th>${monthHeaderCells}<th>Total USD</th><th>TRM</th><th>Total Liquidado COP</th></tr></thead>
    <tbody>${execs.map(e=>{
      const ed=usdData.filter(r=>r['COMERCIAL']===e);
      if(!ed.length) return '';
      const dir=ALL_DATA.find(r=>r['COMERCIAL']===e);
      const monthlyValues = months.map(month => ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])===month).reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0));
      const tot=monthlyValues.reduce((sum, val)=>sum+val,0);
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)">${e}</td>
        <td style="color:var(--text2)">${dir?dir['DIRECTOR']||'—':'—'}</td>
        <td class="td-mono">${ed.length}</td>
        ${monthlyValues.map(val=>`<td class="td-mono td-usd">${val>0?fmtUSD(val):'—'}</td>`).join('')}
        <td class="td-mono td-usd" style="font-weight:700">${fmtUSD(tot)}</td>
        <td class="td-mono" style="color:var(--corp-cyan)">${fmtTRMDisplay(trm)}</td>
        <td class="td-mono td-cop" style="font-weight:700">${fmtCOP(tot*trm)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;color:var(--usd-color)">SUBTOTAL USD</td>
      <td class="td-mono">${usdData.length}</td>
      ${months.map(()=>'<td></td>').join('')}
      <td class="td-mono td-usd" style="font-weight:800">${fmtUSD(totalUSD)}</td>
      <td class="td-mono" style="color:var(--corp-cyan)">${fmtTRMDisplay(trm)}</td>
      <td class="td-mono td-usd" style="font-weight:800;font-size:13px">${fmtCOP(usdLiq)}</td>
    </tr></tbody>
  </table>`;
  
  // Consolidado final
  document.getElementById('tbl-consolidado').innerHTML=`<table>
    <thead><tr><th>Director</th><th>Ejecutivo</th><th>Total COP</th><th>Total USD</th><th>USD Liq. COP</th><th>TOTAL CONSOLIDADO</th></tr></thead>
    <tbody>${dirs.flatMap(d=>{
      const dejecs=[...new Set(ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d).map(r=>r['COMERCIAL']))];
      return dejecs.map((e,ei)=>{
        const ed=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d&&r['COMERCIAL']===e);
        const eCOP=ed.filter(r=>(r['MONEDA 2']||'').trim()==='COP').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
        const eUSD=ed.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
        const total=eCOP+eUSD*trm;
        return `<tr>
          <td style="font-family:var(--font-display);font-weight:700;color:var(--text2)">${ei===0?d:''}</td>
          <td style="color:var(--text)">${e}</td>
          <td class="td-mono td-cop">${fmtCOP(eCOP)}</td>
          <td class="td-mono td-usd">${eUSD>0?fmtUSD(eUSD):'—'}</td>
          <td class="td-mono td-usd">${eUSD>0?fmtCOP(eUSD*trm):'—'}</td>
          <td class="td-mono" style="color:#fff;font-weight:700">${fmtCOP(total)}</td>
        </tr>`;
      });
    }).join('')}
    <tr style="border-top:2px solid var(--corp-blue2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;font-size:13px;color:var(--corp-cyan)">GRAN TOTAL</td>
      <td class="td-mono td-cop" style="font-weight:800">${fmtCOP(totalCOP)}</td>
      <td class="td-mono td-usd" style="font-weight:800">${fmtUSD(totalUSD)}</td>
      <td class="td-mono td-usd" style="font-weight:800">${fmtCOP(usdLiq)}</td>
      <td class="td-mono" style="color:var(--corp-cyan);font-weight:800;font-size:14px">${fmtCOP(totalCOP+usdLiq)}</td>
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
      const estadoClass = ['GANADA','PENDIENTE','PERDIDA','APLAZADO'].includes(estado) ? estado : 'PENDIENTE';
      const rowAttrs = options.clickable && r.__RID
        ? ` class="table-row-action" onclick="openNegocioDetailById('${r.__RID}','${escAttr(sourcePage)}')" title="Abrir detalle del negocio"`
        : '';
      return `<tr${rowAttrs}>
        <td class="td-mono" style="font-size:10px" data-label="Fecha">${escHtml(fecha)}</td>
        <td style="color:var(--text)" data-label="Cliente">${escHtml(cliente)}</td>
        <td style="max-width:130px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${escAttr(producto)}" data-label="Producto">${escHtml(producto)}</td>
        <td style="color:var(--corp-cyan)" data-label="Marca">${escHtml(marca)}</td>
        <td style="font-size:10px" data-label="Línea">${escHtml(linea)}</td>
        <td data-label="Moneda"><span class="badge ${mon==='USD'?'badge-PEDIDA':'badge-PENDIENTE'}">${mon}</span></td>
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
    hideLoadingOverlay();
    alert('Error de autenticación: ' + e.message);
    return;
  }
  updateLoadingStatus('Cargando archivos de SharePoint...');
  try {
    const siteId = await getSiteId();
    const { role, directorGroup } = CURRENT_USER;
    ALL_DATA = [];
    SALES_DATA = [];
    RECORD_SEQ = 0;
    LOADED_FILES_BY_DIR = {};
    LOADED_SALES_BY_SUPPORT = {};

    if(role === 'sales_support') {
      await loadSalesSupportFiles(siteId);
    } else if(role === 'ejecutivo') {
      await loadEjecutivoFile(siteId);
    } else if(role === 'director') {
      const folder = directorGroup.includes('Miller') ? 'Gupo Miller Romero' : 'Grupo ' + directorGroup;
      await loadDirectorFolder(siteId, folder);
    } else {
      // Gerencia — todas las carpetas
      const folders = ['Grupo Juan David Novoa','Grupo Maria Angelica Caballero','Grupo Oscar Beltran','Gupo Miller Romero'];
      for(const f of folders) await loadDirectorFolder(siteId, f);
    }
    finalizeLoad();
    hideLoadingOverlay();
  } catch(e) {
    hideLoadingOverlay();
    console.error(e);
    alert('Error cargando datos: ' + e.message);
  }
}

let _siteId = null;
let _driveId = null;
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

async function loadDirectorFolder(siteId, folderName) {
  const token = await getToken(['Files.Read.All']);
  updateLoadingStatus('Cargando: ' + folderName + '...');
  const folderPath = 'COMERCIAL/FORECAST 2026/' + folderName;
  // Use specific drive if found, otherwise default drive
  const driveBase = _driveId
    ? 'https://graph.microsoft.com/v1.0/drives/' + _driveId + '/root:/' + encodeURIComponent(folderPath) + ':/children?$top=50'
    : 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root:/' + encodeURIComponent(folderPath) + ':/children?$top=50';
  console.log('[FOLDER]', driveBase);
  const r = await fetch(driveBase, { headers: { Authorization: 'Bearer ' + token } });
  const d = await r.json();
  if(!d.value) { console.warn('Sin archivos en', folderName, d); return; }
  const dirName = folderName.replace(/^(Grupo|Gupo)\s+/i,'').trim();
  if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
  for(const item of d.value) {
    if(!item.name.match(/\.xlsx?$/i)) continue;
    if(item.name.startsWith('~$')) continue;
    if(item.name.toLowerCase().includes('base de datos')) continue;
    updateLoadingStatus('Leyendo: ' + item.name);
    const recs = await loadSpFile(item, dirName);
    if(isSalesSupportFile(item.name)) {
      const meta = parseSalesSupportFileName(item.name) || {};
      const supportName = cleanNameSegment(meta.supportName || 'Sales Support');
      if(!LOADED_SALES_BY_SUPPORT[supportName]) LOADED_SALES_BY_SUPPORT[supportName] = [];
      LOADED_SALES_BY_SUPPORT[supportName].push({ name: item.name, dir: dirName, soporta: cleanNameSegment(meta.soportaName) });
      SALES_DATA.push(...recs);
    } else {
      ALL_DATA.push(...recs);
      LOADED_FILES_BY_DIR[dirName].push({ name: item.name });
    }
  }
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

async function loadSalesSupportFiles(siteId) {
  const folders = ['Grupo Juan David Novoa','Grupo Maria Angelica Caballero','Grupo Oscar Beltran','Gupo Miller Romero'];
  const targetName = getSalesSupportTargetName();
  let found = false;
  for(const folder of folders) {
    const folderPath = 'COMERCIAL/FORECAST 2026/' + folder;
    const token = await getToken(['Files.Read.All']);
    try {
      const url = _driveId
        ? 'https://graph.microsoft.com/v1.0/drives/' + _driveId + '/root:/' + encodeURIComponent(folderPath) + ':/children?$top=100'
        : 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root:/' + encodeURIComponent(folderPath) + ':/children?$top=100';
      const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
      const d = await r.json();
      if(!d.value) continue;
      const dirName = folder.replace(/^(Grupo|Gupo)\s+/i,'').trim();
      for(const item of d.value) {
        if(!item.name.match(/\.xlsx?$/i)) continue;
        if(!isSalesSupportFile(item.name)) continue;
        const meta = parseSalesSupportFileName(item.name);
        if(!meta || !namesMatch(meta.supportName, targetName)) continue;
        updateLoadingStatus('Leyendo: ' + item.name);
        const recs = await loadSpFile(item, dirName);
        const supportName = cleanNameSegment(meta.supportName || targetName);
        if(!LOADED_SALES_BY_SUPPORT[supportName]) LOADED_SALES_BY_SUPPORT[supportName] = [];
        LOADED_SALES_BY_SUPPORT[supportName].push({ name: item.name, dir: dirName, soporta: cleanNameSegment(meta.soportaName) });
        SALES_DATA.push(...recs);
        found = true;
      }
    } catch(e) { console.warn('Error leyendo folder sales', folder, e); }
  }
  if(!found) {
    throw new Error('No se encontraron archivos Sales Support para ' + (targetName || 'este usuario') + '.');
  }
}

async function searchExecFileInForecast(siteId, targetName) {
  const email = (CURRENT_USER && CURRENT_USER.email || '').toLowerCase().trim();
  const queries = buildExecSearchQueries(targetName, email);
  if(!queries.length) return null;
  const token = await getToken(['Files.Read.All']);
  const folderPath = 'COMERCIAL/FORECAST 2026';
  for(const q of queries) {
    const url = _driveId
      ? 'https://graph.microsoft.com/v1.0/drives/' + _driveId + '/root:/' + encodeURIComponent(folderPath) + ':/search(q=\'' + encodeURIComponent(q) + '\')?$top=50'
      : 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root:/' + encodeURIComponent(folderPath) + ':/search(q=\'' + encodeURIComponent(q) + '\')?$top=50';
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const d = await r.json();
    const file = findBestExecFile(d.value || [], targetName);
    if(file) return file;
  }
  return null;
}

async function loadEjecutivoFile(siteId) {
  const folders = ['Grupo Juan David Novoa','Grupo Maria Angelica Caballero','Grupo Oscar Beltran','Gupo Miller Romero'];
  const targetName = getExecTargetName();
  let found = false;
  for(const folder of folders) {
    const path = AZURE_CONFIG.driveBase + '/' + folder;
    const token = await getToken(['Files.Read.All']);
    try {
      const r = await fetch(
        _driveId ? 'https://graph.microsoft.com/v1.0/drives/' + _driveId + '/root:/' + encodeURIComponent('COMERCIAL/FORECAST 2026/' + folder) + ':/children?$top=50' : 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root:/' + encodeURIComponent(path) + ':/children?$top=50',
        { headers: { Authorization: 'Bearer ' + token } }
      );
      const d = await r.json();
      if(!d.value) continue;
      const file = findBestExecFile(d.value, targetName);
      if(file) {
        const dirName = folder.replace(/^(Grupo|Gupo)\s+/i,'').trim();
        if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
        const recs = await loadSpFile(file, dirName);
        ALL_DATA.push(...recs);
        LOADED_FILES_BY_DIR[dirName].push({ name: file.name });
        found = true;
        return true;
      }
    } catch(e) { continue; }
  }
  // Fallback: search in FORECAST 2026 if file is nested or renamed
  const fallback = await searchExecFileInForecast(siteId, targetName);
  if(fallback) {
    const dirName = directorFromPath((fallback.parentReference && fallback.parentReference.path) || '') || '';
    if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
    const recs = await loadSpFile(fallback, dirName);
    ALL_DATA.push(...recs);
    LOADED_FILES_BY_DIR[dirName].push({ name: fallback.name });
    return true;
  }
  if(!found) {
    throw new Error('No se encontró el Excel de ' + (targetName || 'este usuario') + '. Verifica el nombre del archivo en FORECAST 2026.');
  }
}

async function loadSpFile(item, dirName) {
  const url = item['@microsoft.graph.downloadUrl'];
  if(!url) return [];
  try {
    const buf = await (await fetch(url)).arrayBuffer();
    const wb  = XLSX.read(buf, { type:'array', cellDates:true });
    const datasetType = isSalesSupportFile(item.name) ? 'sales' : 'forecast';
    const wsName = pickWorksheetName(wb.SheetNames, datasetType);
    const ws  = wb.Sheets[wsName];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    const hdrIdx = findHeaderRowIndex(raw);
    if(hdrIdx<0) return [];
    const hdrs = raw[hdrIdx].map(h=>h ? mapHeaderName(String(h).trim()) : '');
    const recs = [];
    for(let i=hdrIdx+1;i<raw.length;i++) {
      const row=raw[i]; if(!row[1]) continue;
      const rec={};
      hdrs.forEach((h,j)=>{ if(h) rec[h]=row[j]!==undefined?row[j]:null; });
      decorateRecordFromFile(rec, item.name, dirName);
      registerRecord(rec, item.name, wsName, datasetType);
      recs.push(rec);
    }
    return recs;
  } catch(e) { console.warn('Error leyendo', item.name, e); return []; }
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
    divisas:   document.getElementById('tab-divisas'),
    marcas:    document.getElementById('tab-marcas'),
    resumen:   document.getElementById('tab-resumen'),
  };
  // Reset — mostrar todas
  Object.values(tabs).forEach(t => { if(t) t.style.display = ''; });

  if(role === 'sales_support') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.director && (tabs.director.style.display = 'none');
    tabs.ejecutivo && (tabs.ejecutivo.style.display = 'none');
    tabs.divisas  && (tabs.divisas.style.display  = 'none');
    tabs.marcas   && (tabs.marcas.style.display   = 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
    showPage('sales', tabs.sales);
  } else if(role === 'ejecutivo') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.director && (tabs.director.style.display = 'none');
    tabs.sales    && (tabs.sales.style.display    = 'none');
    tabs.divisas  && (tabs.divisas.style.display  = 'none');
    tabs.marcas   && (tabs.marcas.style.display   = 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
    showPage('ejecutivo', tabs.ejecutivo);
  } else if(role === 'director') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.ejecutivo&& (tabs.ejecutivo.style.display= 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
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
window.renderDivisas = renderDivisas;
window.setDivisaEstadoFilter = setDivisaEstadoFilter;
window.showMoreDivisaRows = showMoreDivisaRows;
window.renderDirector = renderDirector;
window.renderEjecutivo = renderEjecutivo;
window.renderSales = renderSales;
window.renderMarcas = renderMarcas;
window.selectEjecutivo = selectEjecutivo;
window.selectSalesSupport = selectSalesSupport;
window.openMarcaLineaDetail = openMarcaLineaDetail;
window.closeMarcaLineaDetail = closeMarcaLineaDetail;
window.setMarcaLineaDetailEstado = setMarcaLineaDetailEstado;
window.setDirectorEstadoFilter = setDirectorEstadoFilter;
window.showMoreEstadoRows = showMoreEstadoRows;
window.openNegocioDetailById = openNegocioDetailById;
window.closeNegocioDetail = closeNegocioDetail;
