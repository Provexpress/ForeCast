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
    classList: { add() {}, remove() {}, contains() { return false; } },
    setAttribute() {},
    removeAttribute() {},
    addEventListener() {},
    querySelectorAll() { return []; }
  };
}

const elements = {};
const getOrCreateElement = id => {
  if (!elements[id]) elements[id] = createElement(id);
  return elements[id];
};

const docMock = {
  documentElement: createElement('documentElement'),
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

const structurePath = path.resolve(__dirname, '..', 'src', 'config', 'estructura-comercial-2026.js');
vm.runInNewContext(fs.readFileSync(structurePath, 'utf8'), context, {
  filename: 'estructura-comercial-2026.js'
});

const structure = context.window.FORECAST_STRUCTURE;

const expectedUsers = [
  {
    email: 'soporte.garantias@provexpress.com.co',
    nombre: 'John Jairo Moreno Quintero'
  },
  {
    email: 'garantias@provexpress.com.co',
    nombre: 'Stefania García Cardenas'
  },
  {
    email: 'oscar.perez@provexpress.com.co',
    nombre: 'Oscar Orlando Perez Mejia'
  }
];

expectedUsers.forEach(u => {
  const role = structure.getRoleByEmail(u.email);
  assert.equal(role, 'glpi_only', `Role for ${u.email} should be glpi_only`);

  const glpiUser = structure.getGlpiUserByEmail(u.email);
  assert.ok(glpiUser, `GLPI user data should exist for ${u.email}`);
  assert.equal(glpiUser.nombre, u.nombre, `Name mismatch for ${u.email}`);

  const displayName = structure.getGlpiUserDisplayNameByEmail(u.email);
  assert.equal(displayName, u.nombre, `Display name mismatch for ${u.email}`);
});

const tabIds = [
  'tab-gerencia', 'tab-director', 'tab-ejecutivo', 'tab-sales',
  'tab-preventa', 'tab-divisas', 'tab-marcas', 'tab-resumen',
  'tab-finanzas', 'tab-programas', 'tab-glpi', 'tab-fondos'
];

tabIds.forEach(getOrCreateElement);

const authPath = path.resolve(__dirname, '..', 'src', 'scripts', 'auth.js');
vm.runInNewContext(fs.readFileSync(authPath, 'utf8'), context, {
  filename: 'auth.js'
});

const mainPath = path.resolve(__dirname, '..', 'src', 'scripts', 'main.js');
vm.runInNewContext(fs.readFileSync(mainPath, 'utf8'), context, {
  filename: 'main.js'
});

const currentUser = {
  email: 'soporte.garantias@provexpress.com.co',
  name: 'John Jairo Moreno Quintero',
  role: 'glpi_only'
};
context.CURRENT_USER = currentUser;
context.window.CURRENT_USER = currentUser;

const applyRoleTabs = context.applyRoleTabs || context.window.applyRoleTabs;
const showPage = context.showPage || context.window.showPage;

applyRoleTabs();

// Verify that tab-glpi is visible and all other tabs are hidden
assert.equal(elements['tab-glpi'].style.display, '', 'tab-glpi should be visible');
tabIds.filter(id => id !== 'tab-glpi').forEach(id => {
  assert.equal(elements[id].style.display, 'none', `${id} should be hidden for glpi_only role`);
});

// Verify showPage guard
let navBlockedWarning = false;
context.console.warn = (msg) => {
  if (msg && msg.includes('acceso denegado a otra vista')) navBlockedWarning = true;
};
showPage('gerencia');
assert.ok(navBlockedWarning, 'showPage should block navigation to non-glpi pages for glpi_only users');

console.log('GLPI User Access: los 3 usuarios de GLPI tienen acceso exclusivo al tablero GLPI y restringen otras vistas.');
