const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const makeEl = () => ({
  value: '', innerHTML: '', style: {}, title: '', dataset: {},
  setAttribute() {}, getAttribute() {},
  addEventListener() {}, removeEventListener() {},
  querySelector() { return null; },
  querySelectorAll() { return []; },
  classList: { add() {}, remove() {}, contains() { return false; } }
});

const elements = new Map();
const getElement = id => {
  if (!elements.has(id)) elements.set(id, makeEl());
  return elements.get(id);
};

[
  'trm-input',
  'sel-director',
  'sel-dir-mes',
  'sel-dir-estado',
  'director-content',
  'donut-dir-est',
  'leg-dir-est',
  'sel-ejecutivo',
  'sel-ej-mes',
  'sel-ej-estado',
  'ejecutivo-content'
].forEach(getElement);

const doc = {
  addEventListener() {},
  removeEventListener() {},
  getElementById: getElement,
  querySelector() { return null; },
  querySelectorAll() { return []; },
  documentElement: { setAttribute() {}, getAttribute() {} },
  body: { classList: { add() {}, remove() {} } }
};

const ctx = {
  window: { addEventListener() {}, removeEventListener() {} },
  document: doc,
  location: { origin: 'https://forecast-provexpress.vercel.app' },
  sessionStorage: { getItem() { return null; }, setItem() {} },
  localStorage: { getItem() { return null; }, setItem() {} },
  CURRENT_USER: { role: 'director', email: 'angelica.caballero@provexpress.com.co', directorGroup: 'Angélica Caballero', group: 2 },
  Intl,
  Math,
  Date
};
ctx.globalThis = ctx.window;
ctx.window.window = ctx.window;
ctx.window.document = doc;
ctx.window.location = ctx.location;
ctx.window.sessionStorage = ctx.sessionStorage;
ctx.window.localStorage = ctx.localStorage;
ctx.window.CURRENT_USER = ctx.CURRENT_USER;

vm.runInNewContext(fs.readFileSync(path.resolve(__dirname, '..', 'src', 'config', 'estructura-comercial-2026.js'), 'utf8'), ctx);
vm.runInNewContext(fs.readFileSync(path.resolve(__dirname, '..', 'src', 'scripts', 'avance-diario.js'), 'utf8'), ctx);
vm.runInNewContext(fs.readFileSync(path.resolve(__dirname, '..', 'src', 'scripts', 'main.js'), 'utf8'), ctx);

// Simular filas de forecast cargadas desde archivos con nombres alternativos / variaciones
const fixtureCode = `
  ALL_DATA = [];
  
  const rec1 = { CLIENTE: 'Cliente 1', ESTADO: 'GANADA', UTILIDAD: 10000000, 'MONEDA 2': 'COP', 'MONTO VENTA CLIENTE': 50000000, 'FECHA DIA/MES/AÑO': '2026-08-15' };
  decorateRecordFromFile(rec1, 'Yurani Vargas.xlsx', 'Angélica Caballero');
  ALL_DATA.push(rec1);

  const rec2 = { CLIENTE: 'Cliente 2', ESTADO: 'GANADA', UTILIDAD: 12000000, 'MONEDA 2': 'COP', 'MONTO VENTA CLIENTE': 60000000, 'FECHA DIA/MES/AÑO': '2026-08-15' };
  decorateRecordFromFile(rec2, 'Yovani Herrera.xlsx', 'Angélica Caballero');
  ALL_DATA.push(rec2);

  const rec3 = { CLIENTE: 'Cliente 3', ESTADO: 'GANADA', UTILIDAD: 15000000, 'MONEDA 2': 'COP', 'MONTO VENTA CLIENTE': 75000000, 'FECHA DIA/MES/AÑO': '2026-08-15' };
  decorateRecordFromFile(rec3, 'Johanna Mojica.xlsx', 'Angélica Caballero');
  ALL_DATA.push(rec3);

  window.TEST_RECS = [rec1, rec2, rec3];

  LOADED_FILES_BY_DIR = {
    'Angelica Caballero': [
      { name: 'Yurani Vargas.xlsx' },
      { name: 'Yovani Herrera.xlsx' },
      { name: 'Johanna Mojica.xlsx' }
    ]
  };

  document.getElementById('sel-director').value = 'Angelica Caballero';
  renderDirector();
`;

vm.runInNewContext(fixtureCode, ctx);

const html = getElement('director-content').innerHTML;

// Verificar que las tarjetas de los 3 ejecutivos muestran sus datos y no "Sin registros aún"
assert.ok(html.includes('Yurany Andrea Vargas'), 'Debe incluir a Yurany Andrea Vargas');
assert.ok(html.includes('Yovanny Herrera'), 'Debe incluir a Yovanny Herrera');
assert.ok(html.includes('Jasbleidy Mójica'), 'Debe incluir a Jasbleidy Mójica');

// Validar que se reconocieron los negocios para cada uno
const recs = ctx.window.TEST_RECS;
assert.ok(ctx.namesMatch(recs[0].COMERCIAL, 'Yurany Andrea Vargas'), 'rec1 debe coincidir con Yurany Andrea Vargas');
assert.ok(ctx.namesMatch(recs[1].COMERCIAL, 'Yovanny Herrera'), 'rec2 debe coincidir con Yovanny Herrera');
assert.ok(ctx.namesMatch(recs[2].COMERCIAL, 'Jasbleidy Mójica'), 'rec3 debe coincidir con Jasbleidy Mójica');

console.log('Aliases Forecast: Cruce de Yurani Vargas, Yovani Herrera y Johanna Mojica validado correctamente.');
