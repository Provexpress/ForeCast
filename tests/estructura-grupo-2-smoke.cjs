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
const expectedGroups = {
  1: [
    'Rosmira Rojas', 'Mario Reyes', 'Wilson Sánchez', 'María Eugenia Cruz',
    'Javier Cortés', 'Rosa Mendoza', 'Mariela Ramírez', 'Jenny Gónzalez',
    'Julieth Galindo'
  ],
  2: [
    'Ángela Torres', 'Yurany Andrea Vargas', 'Alejandra Velásquez',
    'Fernando Quiñonez', 'Jasbleidy Mójica', 'Johanna Jaime', 'Dayana Chala',
    'Yovanny Herrera', 'César Céspedes', 'Daniel Galindo', 'Adriana Cucaita'
  ],
  3: [
    'Gina García', 'Karent Carrillo', 'Lington Linares', 'Angélica Álvarez',
    'Andrés Peña', 'Tatiana Parra', 'Claudia Triana', 'Dilma Cuesta',
    'Juan Martínez', 'Deisy Mogollón'
  ],
  4: [
    'Astrid Jiménez', 'María Paola Briceño', 'Dafne Ruiz', 'Jessica Valencia',
    'Jhonatan Acevedo', 'Camilo Hernández', 'Yeison Urrego', 'Diana Castro'
  ]
};

Object.entries(expectedGroups).forEach(([group, expectedNames]) => {
  const actualNames = structure.getEmailsByGroup(Number(group))
    .map(email => structure.getExecutiveByEmail(email).nombre)
    .sort((a, b) => a.localeCompare(b, 'es'));
  assert.deepEqual(
    Array.from(actualNames, String),
    expectedNames.sort((a, b) => a.localeCompare(b, 'es')),
    `Integrantes incorrectos para el grupo ${group}`
  );
});

const adriana = structure.getExecutiveByEmail('adriana.cucaita@provexpress.com.co');
assert.ok(adriana);
assert.equal(adriana.grupo, 2);
assert.equal(adriana.nombre, 'Adriana Cucaita');
assert.equal(adriana.archivo, 'Adriana Cucaita.xlsx');
assert.equal(structure.getRoleByEmail(adriana.email), 'ejecutivo');
assert.equal(structure.getDirectorNameByGroup(adriana.grupo), 'Angélica Caballero');

const deisy = structure.getExecutiveByEmail('deisy.mogollon@provexpress.com.co');
assert.ok(deisy);
assert.equal(deisy.grupo, 3);
assert.equal(deisy.nombre, 'Deisy Mogollón');
assert.equal(deisy.archivo, 'Deisy Mogollón.xlsx');

const supportNames = emails => emails.map(email => structure.getExecutiveByEmail(email).nombre)
  .sort((a, b) => a.localeCompare(b, 'es'));

assert.deepEqual(
  Array.from(supportNames(structure.getEjecutivosBySupport('soporte.comercial@provexpress.com.co')), String),
  expectedGroups[2].sort((a, b) => a.localeCompare(b, 'es'))
);
assert.deepEqual(
  Array.from(supportNames(structure.getEjecutivosBySupport('soporte.comercial2@provexpress.com.co')), String),
  expectedGroups[3].sort((a, b) => a.localeCompare(b, 'es'))
);
assert.deepEqual(
  Array.from(supportNames(structure.getEjecutivosBySupport('soporte.comercial4@provexpress.com.co')), String),
  ['Yeison Urrego']
);
assert.deepEqual(
  Array.from(supportNames(structure.getEjecutivosBySupport('soporte.comercial3@provexpress.com.co')), String),
  ['Camilo Hernández']
);
assert.deepEqual(
  Array.from(supportNames(structure.getEjecutivosBySupport('soporte.comercial5@provexpress.com.co')), String),
  ['Jessica Valencia', 'María Paola Briceño'].sort((a, b) => a.localeCompare(b, 'es'))
);

assert.equal(structure.getSupportDisplayNameByEmail('soporte.comercial@provexpress.com.co'), 'Karen Cagua');
assert.equal(structure.getSupportDisplayNameByEmail('soporte.comercial2@provexpress.com.co'), 'Alexandra Vargas');
assert.equal(structure.getSupportDisplayNameByEmail('soporte.comercial4@provexpress.com.co'), 'Nury Marcela Vargas');
assert.equal(structure.getSupportDisplayNameByEmail('soporte.comercial3@provexpress.com.co'), 'Janira Alejandra Maldonado');
assert.equal(structure.getSupportDisplayNameByEmail('soporte.comercial5@provexpress.com.co'), 'Johanna Alcocer');

console.log('Estructura comercial: cuatro grupos y cinco Sales Support validados correctamente.');
