const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const RealDate = Date;
const FIXED_NOW = '2026-07-28T12:30:00.000Z';
class FixedDate extends RealDate {
  constructor(...args) {
    super(...(args.length ? args : [FIXED_NOW]));
  }

  static now() {
    return new RealDate(FIXED_NOW).getTime();
  }
}

const elements = new Map();
function getElement(id) {
  if(!elements.has(id)) {
    elements.set(id, {
      id,
      value: '',
      innerHTML: '',
      textContent: '',
      style: {},
      classList: { add() {}, remove() {} },
      setAttribute() {},
      addEventListener() {}
    });
  }
  return elements.get(id);
}

const context = {
  console,
  Date: FixedDate,
  Intl,
  document: {
    body: getElement('body'),
    getElementById: getElement,
    addEventListener() {},
    querySelectorAll() { return []; }
  },
  setTimeout,
  clearTimeout
};
context.window = context;
context.globalThis = context;

const sourcePath = path.resolve(__dirname, '..', 'src', 'scripts', 'glpi.js');
const source = fs.readFileSync(sourcePath, 'utf8');
const marker = '  window.GlpiModule = {';
const instrumented = source.replace(
  marker,
  `  window.__glpiTest = { state, getCurrentPeriod, getPeriods, ensureDefaultPeriod, renderPeriodOptions, getMonthTickets };\n${marker}`
);
assert.notEqual(instrumented, source, 'No se pudo instrumentar el módulo GLPI para la prueba.');

vm.runInNewContext(instrumented, context, { filename: 'glpi.js' });

const api = context.__glpiTest;
assert.equal(api.state.period, 'all');
assert.match(getElement('glpi-period').innerHTML, /value="all" selected/);

api.state.tickets = [{ id:'1', period:'2026-06' }];
api.state.period = '';
api.ensureDefaultPeriod();
api.renderPeriodOptions();

assert.equal(api.state.period, 'all');
assert.match(getElement('glpi-period').innerHTML, /value="all" selected/);
assert.equal(api.getMonthTickets().length, 1);

api.state.period = '2026-06';
api.ensureDefaultPeriod();
assert.equal(api.state.period, '2026-06');
assert.equal(api.getMonthTickets().length, 1);

console.log('GLPI: mes actual de Bogotá disponible aun sin tickets.');
