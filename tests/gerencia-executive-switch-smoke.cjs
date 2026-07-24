const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const elements = new Map();
function getElement(id) {
  if(!elements.has(id)) {
    elements.set(id, {
      id,
      innerHTML: '',
      textContent: '',
      style: {},
      dataset: {},
      classList: { add() {}, remove() {} },
      contains() { return false; },
      appendChild() {}
    });
  }
  return elements.get(id);
}

const documentMock = {
  body: getElement('body'),
  head: getElement('head'),
  getElementById: getElement,
  querySelectorAll() { return []; },
  createElement: tag => getElement(tag),
  addEventListener() {}
};

let roleTabsApplied = 0;

const context = {
  console,
  document: documentMock,
  sessionStorage: { setItem() {}, getItem() { return null; } },
  location: { origin: 'http://127.0.0.1', pathname: '/index.html' },
  applyRoleTabs() { roleTabsApplied += 1; }
};
context.window = context;
context.globalThis = context;

const structurePath = path.resolve(__dirname, '..', 'src', 'config', 'estructura-comercial-2026.js');
const authPath = path.resolve(__dirname, '..', 'src', 'scripts', 'auth.js');
const structureSource = fs.readFileSync(structurePath, 'utf8');
const authSource = fs.readFileSync(authPath, 'utf8');
const fixture = `
  CURRENT_USER = {
    email: 'especialista.preventa@provexpress.com.co',
    name: 'Especialista Preventa',
    role: 'gerencia'
  };
  renderViewPanelOptions();
`;

vm.runInNewContext(`${structureSource}\n${authSource}\n${fixture}`, context, { filename: 'auth.js' });

const optionsHtml = getElement('view-panel-options').innerHTML;
assert.match(optionsHtml, /Vista como comercial/);
assert.match(optionsHtml, /Rosmira Rojas/);
assert.match(optionsHtml, /Rafael Novoa/);
assert.ok((optionsHtml.match(/class="view-opt-btn view-opt-executive"/g) || []).length > 20);

const executiveButton = {
  dataset: {
    executiveName: 'Rosmira Rojas',
    directorName: 'Rafael Novoa'
  },
  classList: { add() {}, remove() {} }
};
context.switchExecutiveView(executiveButton);

assert.equal(context.CURRENT_USER.role, 'ejecutivo');
assert.equal(context.CURRENT_USER.name, 'Rosmira Rojas');
assert.equal(context.CURRENT_USER.directorGroup, 'Rafael Novoa');
assert.equal(roleTabsApplied, 1);

console.log('Gerencia: cambio a vista comercial validado correctamente.');
