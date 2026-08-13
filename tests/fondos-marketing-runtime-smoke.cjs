const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const elements = new Map();

class FakeElement {
  constructor(id = '') {
    this.id = id;
    this._innerHTML = '';
    this.style = {};
  }

  set innerHTML(value) {
    this._innerHTML = String(value);
    for(const match of this._innerHTML.matchAll(/id="([^"]+)"/g)) {
      if(!elements.has(match[1])) elements.set(match[1], new FakeElement(match[1]));
    }
  }

  get innerHTML() {
    return this._innerHTML;
  }

  addEventListener() {}
  getAttribute() { return null; }
  querySelector(selector) {
    if(selector === '.fondos-app-wrapper' && this._innerHTML.includes('class="fondos-app-wrapper"')) return {};
    return null;
  }
  querySelectorAll() { return []; }
}

elements.set('page-fondos', new FakeElement('page-fondos'));

const document = {
  getElementById(id) {
    return elements.get(id) || null;
  }
};

const context = {
  console,
  document,
  CURRENT_USER: { role: 'gerencia', email: 'especialista.preventa@provexpress.com.co' },
  Intl,
  JSON,
  Math,
  Number,
  Date,
  setTimeout,
  clearTimeout
};
context.window = context;

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'src', 'scripts', 'fondos-marketing.js'), 'utf8');
vm.runInNewContext(source, context, { filename: 'fondos-marketing.js' });

assert.ok(context.FondosMarketingModule, 'El módulo debe registrarse en window');
assert.equal(context.FondosMarketingModule.canAccess(), true);
assert.doesNotThrow(() => context.FondosMarketingModule.init());
assert.match(elements.get('page-fondos').innerHTML, /Fondos de Mercadeo \(MDF\)/);
assert.match(elements.get('fondos-kpis-container').innerHTML, /Ingresos Totales/);
assert.match(elements.get('fondos-summary-tbody').innerHTML, /TOTAL CONSOLIDADO/);
assert.match(elements.get('fondos-brand-detail-container').innerHTML, /HP/);

for(const id of [...elements.keys()]) {
  if(id !== 'page-fondos') elements.delete(id);
}
elements.get('page-fondos').innerHTML = '';
assert.doesNotThrow(() => context.FondosMarketingModule.init());
assert.match(elements.get('page-fondos').innerHTML, /Fondos de Mercadeo \(MDF\)/);
assert.match(elements.get('fondos-kpis-container').innerHTML, /Ingresos Totales/);

console.log('Fondos Marketing: renderizado inicial y reconstrucción completa validados.');
