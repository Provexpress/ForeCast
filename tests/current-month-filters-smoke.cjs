const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const RealDate = Date;
const FIXED_NOW = '2026-08-01T02:30:00.000Z';
class FixedDate extends RealDate {
  constructor(...args) {
    super(...(args.length ? args : [FIXED_NOW]));
  }

  static now() {
    return new RealDate(FIXED_NOW).getTime();
  }
}

function createElement(id = '') {
  const attributes = {};
  return {
    id,
    value: '',
    innerHTML: '',
    textContent: '',
    title: '',
    min: '',
    max: '',
    style: {},
    dataset: {},
    disabled: false,
    parentElement: null,
    classList: {
      add() {},
      remove() {},
      contains() { return false; }
    },
    addEventListener() {},
    removeEventListener() {},
    setAttribute(name, value) { attributes[name] = String(value); },
    getAttribute(name) { return attributes[name] || null; },
    removeAttribute(name) { delete attributes[name]; },
    querySelectorAll() { return []; },
    querySelector() { return null; },
    appendChild() {},
    remove() {}
  };
}

const elements = new Map();
const getElement = id => {
  if(!elements.has(id)) elements.set(id, createElement(id));
  return elements.get(id);
};

[
  'sel-gerencia-mes',
  'sel-dir-mes',
  'sel-ej-mes',
  'sel-sales-mes',
  'sel-preventa-mes',
  'finance-period',
  'finance-end-date',
  'finance-goal',
  'theme-toggle-btn',
  'trm-input'
].forEach(getElement);

const documentMock = {
  body: createElement('body'),
  head: createElement('head'),
  documentElement: createElement('html'),
  getElementById: getElement,
  querySelectorAll() { return []; },
  querySelector() { return null; },
  createElement: tag => createElement(tag),
  addEventListener() {},
  removeEventListener() {},
  dispatchEvent() {}
};

const storageMock = {
  getItem() { return null; },
  setItem() {},
  removeItem() {}
};

const context = {
  console,
  document: documentMock,
  localStorage: storageMock,
  sessionStorage: storageMock,
  navigator: {},
  location: { origin: 'http://127.0.0.1', pathname: '/index.html' },
  fetch: async () => ({ ok: false, json: async () => ({}) }),
  alert() {},
  confirm() { return true; },
  setTimeout,
  clearTimeout,
  Intl,
  Date: FixedDate,
  Math,
  Map,
  Set,
  URL,
  Blob,
  CURRENT_USER: null,
  FORECAST_STRUCTURE: {},
  addEventListener() {},
  removeEventListener() {},
  scrollTo() {},
  innerWidth: 1440,
  innerHeight: 900
};
context.window = context;
context.globalThis = context;

const sourcePath = path.resolve(__dirname, '..', 'src', 'scripts', 'main.js');
const source = fs.readFileSync(sourcePath, 'utf8');
const fixture = `
  ALL_DATA = [];
  SALES_DATA = [];
  PREVENTA_DATA = [];
  refreshForecastMonthFilters();
  window.__initialValues = [
    document.getElementById('sel-gerencia-mes').value,
    document.getElementById('sel-dir-mes').value,
    document.getElementById('sel-ej-mes').value,
    document.getElementById('sel-sales-mes').value,
    document.getElementById('sel-preventa-mes').value
  ];
  window.__initialMarkup = [
    document.getElementById('sel-gerencia-mes').innerHTML,
    document.getElementById('sel-dir-mes').innerHTML,
    document.getElementById('sel-ej-mes').innerHTML,
    document.getElementById('sel-sales-mes').innerHTML,
    document.getElementById('sel-preventa-mes').innerHTML
  ];

  ALL_DATA = [
    { 'FECHA DIA/MES/AÑO':'2026-06-10' },
    {
      'FECHA DIA/MES/AÑO':'2026-07-10',
      'MONEDA 2': 2,
      'MONTO VENTA CLIENTE': 1250000,
      'DIRECTOR': 'Angélica Caballero',
      'COMERCIAL': 'Adriana Cucaita',
      'ESTADO': 'PENDIENTE'
    }
  ];
  SALES_DATA = [{ 'FECHA DIA/MES/AÑO':'2026-06-11' }];
  PREVENTA_DATA = [{ 'FECHA DIA/MES/AÑO':'2026-06-12' }];
  renderGerencia();
  window.__gerenciaCurrentMonth = {
    chart: document.getElementById('evo-dir-chart').innerHTML,
    kpis: document.getElementById('kpi-gerencia').innerHTML,
    normalizedCurrency: getRowCurrency(ALL_DATA[1]),
    valueCop: toCOP(ALL_DATA[1])
  };
  document.getElementById('sel-director').value = 'Angélica Caballero';
  document.getElementById('sel-dir-estado').value = '';
  renderDirector();
  window.__angelicaDirectorKpis = document.getElementById('director-content').innerHTML;

  GERENCIA_CROSSFILTERS.mes = '2026-06';
  GERENCIA_MONTH_INITIALIZED = true;
  ['sel-dir-mes','sel-ej-mes','sel-sales-mes','sel-preventa-mes'].forEach(id => {
    document.getElementById(id).value = '2026-06';
  });
  refreshForecastMonthFilters();
  window.__manualValues = [
    document.getElementById('sel-gerencia-mes').value,
    document.getElementById('sel-dir-mes').value,
    document.getElementById('sel-ej-mes').value,
    document.getElementById('sel-sales-mes').value,
    document.getElementById('sel-preventa-mes').value
  ];

  GERENCIA_CROSSFILTERS.mes = '';
  ['sel-dir-mes','sel-ej-mes','sel-sales-mes','sel-preventa-mes'].forEach(id => {
    document.getElementById(id).value = '';
  });
  refreshForecastMonthFilters();
  window.__allMonthsValues = [
    document.getElementById('sel-gerencia-mes').value,
    document.getElementById('sel-dir-mes').value,
    document.getElementById('sel-ej-mes').value,
    document.getElementById('sel-sales-mes').value,
    document.getElementById('sel-preventa-mes').value
  ];

  FINANCE_STATE = { isLoading:false, data:null, error:null, period:'', endDate:'' };
  setFinanceInputs();
  window.__financeDefault = {
    period: document.getElementById('finance-period').value,
    endDate: document.getElementById('finance-end-date').value
  };
  FINANCE_STATE = { isLoading:false, data:null, error:null, period:'2026-06', endDate:'' };
  setFinanceInputs();
  window.__financeManual = document.getElementById('finance-period').value;

  FORECAST_STRUCTURE = {
    getEjecutivosBySupport(email) {
      return email === 'soporte.comercial@provexpress.com.co'
        ? ['angela.torres@provexpress.com.co']
        : [];
    },
    getExecutiveMatchNamesByEmail(email) {
      return email === 'angela.torres@provexpress.com.co' ? ['Ángela Torres'] : [];
    },
    getSupportDisplayNameByEmail() { return 'Karen Cagua'; }
  };
  CURRENT_USER = {
    email: 'soporte.comercial@provexpress.com.co',
    role: 'sales_support',
    supportScope: 'unit',
    directorGroup: 'Angélica Caballero'
  };
  ALL_DATA = [
    { COMERCIAL:'Ángela Torres' },
    { COMERCIAL:'Gina García' }
  ];
  SALES_DATA = [
    { 'SALES SUPPORT':'Karen Cagua', SOPORTA:'Ángela Torres' },
    { 'SALES SUPPORT':'Karen Cagua', SOPORTA:'Gina García' }
  ];
  window.__unitSupportScope = {
    forecast: getVisibleData().map(row => row.COMERCIAL),
    sales: getVisibleSalesData().map(row => row.SOPORTA),
    badge: buildSalesSupportScopeBadge()
  };
`;

vm.runInNewContext(`${source}\n${fixture}`, context, { filename: 'main.js' });

assert.deepEqual([...context.__initialValues], Array(5).fill('2026-07'));
context.__initialMarkup.forEach(markup => {
  assert.match(markup, /value="2026-07"/);
  assert.match(markup, /julio 2026/i);
});
assert.match(context.__gerenciaCurrentMonth.chart, />Jul</);
assert.doesNotMatch(context.__gerenciaCurrentMonth.chart, />Jun</);
assert.match(context.__gerenciaCurrentMonth.kpis, /kpi-val">\s*\$\s*1\.3 M/);
assert.equal(context.__gerenciaCurrentMonth.normalizedCurrency, 'COP');
assert.equal(context.__gerenciaCurrentMonth.valueCop, 1250000);
assert.match(context.__angelicaDirectorKpis, /kpi-val">\s*\$\s*1\.3 M/);
assert.deepEqual([...context.__manualValues], Array(5).fill('2026-06'));
assert.deepEqual([...context.__allMonthsValues], Array(5).fill(''));
assert.equal(context.__financeDefault.period, '2026-07');
assert.equal(context.__financeDefault.endDate, '2026-07-31');
assert.equal(context.__financeManual, '2026-06');
assert.deepEqual([...context.__unitSupportScope.forecast], ['Ángela Torres']);
assert.deepEqual([...context.__unitSupportScope.sales], ['Ángela Torres']);
assert.match(context.__unitSupportScope.badge, /UNIDAD ANGÉLICA CABALLERO/);

console.log('Filtros mensuales: mes actual Bogotá y selecciones manuales validados.');
