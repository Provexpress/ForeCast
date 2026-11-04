const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createElement(id = '', tag = 'div') {
  return {
    id,
    tagName: tag.toUpperCase(),
    value: '',
    innerHTML: '',
    textContent: '',
    style: {},
    dataset: {},
    classList: { add() {}, remove() {}, contains() { return false; } },
    setAttribute() {},
    removeAttribute() {},
    addEventListener() {},
    removeEventListener() {},
    querySelectorAll() { return []; },
    querySelector() { return null; },
    appendChild() {},
    remove() {},
    contains() { return false; }
  };
}

const elements = {};
const getOrCreateElement = id => {
  if (!elements[id]) elements[id] = createElement(id);
  return elements[id];
};

const docMock = {
  documentElement: createElement('documentElement'),
  createElement: (tag) => createElement('', tag),
  getElementById: getOrCreateElement,
  querySelector: () => null,
  querySelectorAll: () => [],
  addEventListener: () => {}
};

const context = {
  window: {
    addEventListener: () => {}
  },
  document: docMock,
  console: console
};
context.globalThis = context.window;
context.window.window = context.window;
context.window.document = docMock;

context.window.localStorage = {
  getItem: () => null,
  setItem: () => {},
  removeItem: () => {}
};

context.window.sessionStorage = {
  getItem: () => null,
  setItem: () => {},
  removeItem: () => {}
};

context.window.location = {
  origin: 'https://forecast.local',
  pathname: '/'
};

// 1. Cargar estructura comercial
const structurePath = path.resolve(__dirname, '..', 'src', 'config', 'estructura-comercial-2026.js');
vm.runInNewContext(fs.readFileSync(structurePath, 'utf8'), context, {
  filename: 'estructura-comercial-2026.js'
});

const structure = context.window.FORECAST_STRUCTURE;

// 2. Validar rol de Jorge Sepúlveda
const jorgeEmail = 'jorge.sepulveda@provexpress.com.co';
const role = structure.getRoleByEmail(jorgeEmail);
assert.equal(role, 'preventa', `Role for ${jorgeEmail} should be preventa`);

// Validar que gerencia (incluyendo especialista.preventa) conserve rol gerencia
assert.equal(structure.getRoleByEmail('especialista.preventa@provexpress.com.co'), 'gerencia');
assert.equal(structure.getRoleByEmail('preventa.software@provexpress.com.co'), 'gerencia');

// 3. Validar datos mapeados de preventa
const preventaUser = structure.getPreventaByEmail(jorgeEmail);
assert.ok(preventaUser, `Preventa user data should exist for ${jorgeEmail}`);
assert.equal(preventaUser.nombre, 'Jorge Sepúlveda Garavito');
assert.equal(preventaUser.archivo, 'Jorge Sepulveda Garavito.xlsx');

assert.equal(structure.getPreventaDisplayNameByEmail(jorgeEmail), 'Jorge Sepúlveda Garavito');
assert.equal(structure.getPreventaFileByEmail(jorgeEmail), 'Jorge Sepulveda Garavito.xlsx');
assert.equal(structure.getPreventaFileBaseByEmail(jorgeEmail), 'Jorge Sepulveda Garavito');
assert.equal(structure.getPreventaEmailByName('Jorge Sepulveda Garavito'), jorgeEmail);
assert.equal(structure.getPreventaEmailByName('Jorge Sepulveda Garavito.xlsx'), jorgeEmail);
assert.equal(structure.canonicalizePreventaName('Jorge Sepulveda'), 'Jorge Sepúlveda Garavito');

// 4. Cargar scripts de auth y main
const tabIds = [
  'tab-gerencia', 'tab-director', 'tab-ejecutivo', 'tab-sales',
  'tab-preventa', 'tab-divisas', 'tab-marcas', 'tab-resumen',
  'tab-finanzas', 'tab-programas', 'tab-glpi', 'tab-fondos'
];
tabIds.forEach(getOrCreateElement);

// Elementos de la vista Preventa
getOrCreateElement('preventa-content');
getOrCreateElement('preventa-grid');
getOrCreateElement('sel-preventa');
getOrCreateElement('sel-preventa-mes');
getOrCreateElement('sel-preventa-estado');
getOrCreateElement('page-preventa');
getOrCreateElement('page-gerencia');
getOrCreateElement('page-negocio');

const authPath = path.resolve(__dirname, '..', 'src', 'scripts', 'auth.js');
vm.runInNewContext(fs.readFileSync(authPath, 'utf8'), context, {
  filename: 'auth.js'
});

const mainPath = path.resolve(__dirname, '..', 'src', 'scripts', 'main.js');
vm.runInNewContext(fs.readFileSync(mainPath, 'utf8'), context, {
  filename: 'main.js'
});

// 5. Validar auth helpers
const configuredName = context.getConfiguredUserName
  ? context.getConfiguredUserName(jorgeEmail, 'preventa')
  : structure.getPreventaDisplayNameByEmail(jorgeEmail);
assert.equal(configuredName, 'Jorge Sepúlveda Garavito');

// 6. Validar sesión como Preventa
const currentUser = {
  email: jorgeEmail,
  name: 'Jorge Sepúlveda Garavito',
  role: 'preventa'
};
context.CURRENT_USER = currentUser;
context.window.CURRENT_USER = currentUser;

const applyRoleTabs = context.applyRoleTabs || context.window.applyRoleTabs;
const showPage = context.showPage || context.window.showPage;
const namesMatch = context.namesMatch || context.window.namesMatch;

applyRoleTabs();

// Verificar que tab-preventa esté visible y todas las demás pestañas ocultas
assert.equal(elements['tab-preventa'].style.display, '', 'tab-preventa should be visible');
tabIds.filter(id => id !== 'tab-preventa').forEach(id => {
  assert.equal(elements[id].style.display, 'none', `${id} should be hidden for preventa role`);
});

// Verificar que showPage bloquee navegación a gerencia u otras páginas no permitidas
let navBlockedWarning = false;
context.console.warn = (msg) => {
  if (msg && msg.includes('acceso denegado a otra vista para rol preventa')) navBlockedWarning = true;
};
showPage('gerencia');
assert.ok(navBlockedWarning, 'showPage should block navigation to non-preventa pages for preventa users');

// Verificar namesMatch entre variaciones de nombres de Jorge
assert.ok(namesMatch('Jorge Sepúlveda Garavito', 'Jorge Sepulveda Garavito.xlsx'));
assert.ok(namesMatch('Jorge Sepulveda', 'Jorge Sepulveda Garavito'));

console.log('Preventa User Access: Jorge Sepúlveda Garavito configurado correctamente con acceso exclusivo a Preventa y archivo mapeado.');
