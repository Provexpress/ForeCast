const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const RealDate = Date;
const FIXED_NOW = '2026-07-27T17:00:00.000Z';
class FixedDate extends RealDate {
  constructor(...args) {
    super(...(args.length ? args : [FIXED_NOW]));
  }

  static now() {
    return new RealDate(FIXED_NOW).getTime();
  }
}

function createElement(id = '') {
  return {
    id,
    value: '',
    innerHTML: '',
    textContent: '',
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
    setAttribute() {},
    removeAttribute() {},
    querySelectorAll() { return []; },
    querySelector() { return null; },
    appendChild() {},
    remove() {},
    getBoundingClientRect() {
      return { left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 };
    }
  };
}

const elements = new Map();
const getElement = id => {
  if(!elements.has(id)) elements.set(id, createElement(id));
  return elements.get(id);
};

[
  'trm-input',
  'sel-ejecutivo',
  'sel-ej-mes',
  'sel-ej-estado',
  'persona-grid',
  'ejecutivo-content',
  'bar-ej-lineas',
  'donut-ej-est',
  'leg-ej-est'
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
  removeEventListener() {}
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
  CURRENT_USER: { role: 'gerencia', email: 'test@provexpress.com.co' },
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
  ALL_DATA = [{
    COMERCIAL: 'Juan David Martínez Pérez',
    DIRECTOR: 'Angélica Caballero',
    CLIENTE: 'Cliente prueba',
    ESTADO: 'GANADA',
    UTILIDAD: 5000000,
    'MONEDA 2': 'COP',
    'MONTO VENTA CLIENTE': 50000000,
    'FECHA DIA/MES/AÑO': '2026-07-10',
    'LINEA DE PRODUCTO': 'Tecnología'
  }];
  LOADED_FILES_BY_DIR = {
    'Angélica Caballero': [{ name: 'Juan David Martínez.xlsx' }]
  };
  renderEjecutivo();
  window.__contentWithData = document.getElementById('ejecutivo-content').innerHTML;

  ALL_DATA[0]['FECHA DIA/MES/AÑO'] = '2026-06-10';
  ALL_DATA[0].UTILIDAD = 9000000;
  renderEjecutivo();
  window.__contentWithHistoricalData = document.getElementById('ejecutivo-content').innerHTML;
  window.__cardWithHistoricalData = document.getElementById('persona-grid').innerHTML;

  ALL_DATA = [];
  LOADED_FILES_BY_DIR = {
    'Angélica Caballero': [{ name: 'Adriana Cucaita.xlsx' }]
  };
  document.getElementById('sel-ejecutivo').value = 'Adriana Cucaita';
  renderEjecutivo();
  window.__contentWithoutData = document.getElementById('ejecutivo-content').innerHTML;
`;

vm.runInNewContext(`${source}\n${fixture}`, context, { filename: 'main.js' });

assert.match(context.__contentWithData, /Juan David Martínez Pérez/);
assert.match(context.__contentWithData, /quota-summary--executive/);
assert.match(context.__contentWithData, /Mi cuota mensual/i);
assert.match(context.__contentWithData, /14[.\s\u00a0]000[.\s\u00a0]000/);
assert.match(context.__contentWithData, /Utilidad lograda/i);
assert.match(context.__contentWithData, /5[.\s\u00a0]000[.\s\u00a0]000/);
assert.match(context.__contentWithData, /role="progressbar"/);
assert.match(context.__contentWithHistoricalData, /Utilidad lograda/i);
assert.match(context.__contentWithHistoricalData, />\$\s*0</);
assert.match(context.__cardWithHistoricalData, /persona-card selected no-data/);
assert.doesNotMatch(context.__cardWithHistoricalData, /\$9(?:[.,]0)?M/);
assert.match(context.__contentWithoutData, /Adriana Cucaita/);
assert.match(context.__contentWithoutData, /18[.\s\u00a0]000[.\s\u00a0]000/);
assert.match(context.__contentWithoutData, /Utilidad lograda/i);
assert.match(context.__contentWithoutData, />\$\s*0</);
assert.equal((getElement('persona-grid').innerHTML.match(/persona-card/g) || []).length, 1);

console.log('Vista Ejecutivo: cuota y utilidad renderizadas correctamente.');
