const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const context = { window: {} };
context.globalThis = context.window;

const structurePath = path.resolve(__dirname, '..', 'src', 'config', 'estructura-comercial-2026.js');
vm.runInNewContext(fs.readFileSync(structurePath, 'utf8'), context, {
  filename: 'estructura-comercial-2026.js'
});

const structure = context.window.FORECAST_STRUCTURE;
const expectedTeam = [
  'Adriana Cucaita',
  'Alejandra Velásquez',
  'Dayana Chala',
  'Daniel Galindo',
  'Fernando Quiñonez',
  'Jasbleidy Mójica',
  'Johanna Jaime',
  'Yovanny Herrera',
  'Yurany Andrea Vargas',
  'Ángela Torres',
  'César Céspedes'
].sort((a, b) => a.localeCompare(b, 'es'));

const groupTwo = structure.getEmailsByGroup(2)
  .map(email => structure.getExecutiveByEmail(email))
  .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));

assert.deepEqual(
  Array.from(groupTwo, executive => String(executive.nombre)),
  expectedTeam
);

const adriana = structure.getExecutiveByEmail('adriana.cucaita@provexpress.com.co');
assert.ok(adriana);
assert.equal(adriana.grupo, 2);
assert.equal(adriana.nombre, 'Adriana Cucaita');
assert.equal(adriana.archivo, 'Adriana Cucaita.xlsx');
assert.equal(structure.getRoleByEmail(adriana.email), 'ejecutivo');
assert.equal(structure.getDirectorNameByGroup(adriana.grupo), 'Angélica Caballero');

console.log('Grupo 2: Adriana Cucaita y los 11 integrantes validados correctamente.');
