const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

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
    classList: { add() {}, remove() {}, contains() { return false; } },
    addEventListener() {},
    removeEventListener() {},
    setAttribute() {},
    removeAttribute() {},
    querySelectorAll() { return []; },
    querySelector() { return null; },
    appendChild() {},
    remove() {},
    getBoundingClientRect() { return { left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 }; }
  };
}

const elements = new Map();
const getElement = id => {
  if(!elements.has(id)) elements.set(id, createElement(id));
  return elements.get(id);
};

[
  'trm-input',
  'sel-director',
  'sel-dir-mes',
  'sel-dir-estado',
  'director-content',
  'donut-dir-est',
  'leg-dir-est'
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

const storageMock = { getItem() { return null; }, setItem() {}, removeItem() {} };

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
  Date,
  Math,
  Map,
  Set,
  URL,
  Blob,
  CURRENT_USER: { role: 'director', email: 'angelica.caballero@provexpress.com.co', directorGroup: 'Angélica Caballero', group: 2 },
  FORECAST_STRUCTURE: {},
  addEventListener() {},
  removeEventListener() {},
  scrollTo() {},
  innerWidth: 1440,
  innerHeight: 900
};
context.window = context;
context.globalThis = context;

const estCode = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'config', 'estructura-comercial-2026.js'), 'utf8');
const avanceCode = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'scripts', 'avance-diario.js'), 'utf8');
const mainCode = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'scripts', 'main.js'), 'utf8');

const fixture = `
  ALL_DATA = [{
    COMERCIAL: 'Ángela Torres',
    DIRECTOR: 'Angélica Caballero',
    CLIENTE: 'Cliente prueba',
    ESTADO: 'GANADA',
    UTILIDAD: 5000000,
    'MONEDA 2': 'COP',
    'MONTO VENTA CLIENTE': 50000000,
    'FECHA DIA/MES/AÑO': '2026-08-10',
    'LINEA DE PRODUCTO': 'Tecnología'
  }];
  LOADED_FILES_BY_DIR = {
    'Angelica Caballero': [{ name: 'Ángela Torres.xlsx' }]
  };
  document.getElementById('sel-director').value = 'Angelica Caballero';
  renderDirector();
`;

vm.runInNewContext(estCode + '\n' + avanceCode + '\n' + mainCode + '\n' + fixture, context);

const html = getElement('director-content').innerHTML;

assert.ok(html.includes('11 comerciales asignados'), 'Debe indicar 11 comerciales asignados en el resumen');
assert.ok(html.includes('11 Comerciales Evaluados'), 'Debe indicar 11 Comerciales Evaluados');
assert.ok(html.includes('11 EJECUTIVOS'), 'Debe indicar 11 EJECUTIVOS en la sección de equipo');

const expectedComerciales = [
  'Ángela Torres', 'Yurany Andrea Vargas', 'Alejandra Velásquez',
  'Fernando Quiñonez', 'Jasbleidy Mójica', 'Johanna Jaime', 'Dayana Chala',
  'Yovanny Herrera', 'César Céspedes', 'Daniel Galindo', 'Adriana Cucaita'
];

expectedComerciales.forEach(name => {
  assert.ok(html.includes(name), 'Debe incluir en la tabla y tarjetas a ' + name);
});

console.log('Director Grupo 2 (Angélica Caballero): Todos los 11 comerciales renderizados correctamente en vista de director.');
