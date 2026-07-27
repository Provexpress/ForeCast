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
    classList: { add() {}, remove() {}, contains() { return false; } },
    setAttribute() {},
    addEventListener() {},
    querySelectorAll() { return []; }
  };
}

const elements = new Map();
const getElement = id => {
  if(!elements.has(id)) elements.set(id, createElement(id));
  return elements.get(id);
};

const context = {
  console,
  Date: FixedDate,
  Intl,
  document: {
    body: createElement('body'),
    getElementById: getElement,
    addEventListener() {},
    querySelectorAll() { return []; }
  },
  setTimeout,
  clearTimeout,
  FORECAST_STRUCTURE: {}
};
context.window = context;
context.globalThis = context;

const sourcePath = path.resolve(__dirname, '..', 'src', 'scripts', 'glpi.js');
const source = fs.readFileSync(sourcePath, 'utf8');
const marker = '  window.GlpiModule = {';
const exposed = [
  'state',
  'getContextMetrics',
  'getFilteredTickets',
  'renderKpis',
  'renderFilterOptions',
  'renderGroups',
  'renderTable',
  'summarizeBy',
  'setGroup',
  'setRequester',
  'setStatus',
  'toggleStatus',
  'getTicketTimeInfo',
  'normalizeTicketRow'
].join(', ');
const instrumented = source.replace(
  marker,
  `  window.__glpiCrossfilterTest = { ${exposed} };\n${marker}`
);
assert.notEqual(instrumented, source, 'No se pudo instrumentar GLPI.');

vm.runInNewContext(instrumented, context, { filename:'glpi.js' });
const api = context.__glpiCrossfilterTest;

const normalizedOpenTicket = api.normalizeTicketRow({
  ID:8,
  'Fecha de Apertura':'2026-07-20',
  Estado:'En curso',
  Solicitante:'Persona Prueba',
  'Días Abierto':1
}, 0);
assert.equal(normalizedOpenTicket.daysOpen, 7, 'Los días abiertos deben actualizarse hasta hoy en Bogotá.');

const normalizedClosedTicket = api.normalizeTicketRow({
  ID:9,
  'Fecha de Apertura':'2026-07-01',
  'Fecha de solución':'2026-07-09',
  Estado:'Cerrado',
  Solicitante:'Persona Prueba'
}, 1);
assert.equal(normalizedClosedTicket.daysSolution, 8, 'El cierre debe medir días entre apertura y solución.');

const closedWithoutDate = api.normalizeTicketRow({
  ID:10,
  'Fecha de Apertura':'2026-07-01',
  Estado:'Cerrado',
  Solicitante:'Persona Prueba'
}, 2);
assert.equal(closedWithoutDate.daysSolution, null, 'Un cierre sin fecha no debe inventar duración.');

function ticket(id, statusKey, group, requester, times = {}) {
  return {
    id:String(id),
    title:`Ticket ${id}`,
    status:statusKey === 'closed' ? 'Cerrado' : (statusKey === 'new' ? 'Nuevo' : 'En curso'),
    statusKey,
    openDate:'2026-07-01',
    solutionDate:statusKey === 'closed' ? '2026-07-10' : '',
    period:'2026-07',
    technicians:['Técnico Uno'],
    category:'Soporte > Prueba',
    categoryGroup:'Soporte',
    directorGroup:group,
    requester,
    description:'',
    daysOpen:times.daysOpen ?? null,
    daysSolution:times.daysSolution ?? null
  };
}

api.state.loaded = true;
api.state.period = '2026-07';
api.state.tickets = [
  ticket(1, 'new', 'Grupo A', 'Ana', { daysOpen:2 }),
  ticket(2, 'open', 'Grupo A', 'Ana', { daysOpen:10 }),
  ticket(3, 'closed', 'Grupo A', 'Ana', { daysSolution:6 }),
  ticket(4, 'open', 'Grupo A', 'Beto', { daysOpen:4 }),
  ticket(5, 'closed', 'Grupo A', 'Beto', { daysSolution:2 }),
  ticket(6, 'open', 'Grupo B', 'Carla', { daysOpen:20 }),
  ticket(7, 'closed', 'Grupo B', 'Carla', { daysSolution:10 })
];

api.setGroup('Grupo A');
let metrics = api.getContextMetrics();
assert.equal(metrics.context.length, 5);
assert.equal(metrics.pendingTickets.length, 3);
assert.equal(metrics.newTickets.length, 1);
assert.equal(metrics.openTickets.length, 2);
assert.equal(metrics.averageOpenDays, 16 / 3);
assert.equal(metrics.medianOpenDays, 4);
assert.equal(metrics.maxOpenDays, 10);
assert.deepEqual([...metrics.aboveAverageTickets.map(item => item.id)], ['2']);
assert.equal(metrics.closedTickets.length, 2);
assert.equal(metrics.averageClosedDays, 4);
assert.equal(metrics.medianClosedDays, 4);
assert.equal(metrics.maxClosedDays, 6);

api.renderKpis();
const groupCards = getElement('glpi-kpis').innerHTML;
assert.match(groupCards, /Tickets del contexto[\s\S]*>5</);
assert.match(groupCards, /Pendientes[\s\S]*>3</);
assert.match(groupCards, /Antigüedad abierta[\s\S]*>5,3</);
assert.match(groupCards, /Pendientes sobre promedio[\s\S]*>1</);
assert.match(groupCards, /Cerrados[\s\S]*>2</);
assert.match(groupCards, /Prom\. 4 días para cerrar/);
assert.match(groupCards, /Tiempo de cierre[\s\S]*>4</);

api.setRequester('Ana');
metrics = api.getContextMetrics();
assert.equal(metrics.context.length, 3);
assert.equal(metrics.pendingTickets.length, 2);
assert.equal(metrics.averageOpenDays, 6);
assert.equal(metrics.averageClosedDays, 6);
assert.deepEqual([...metrics.aboveAverageTickets.map(item => item.id)], ['2']);
assert.match(getElement('glpi-context-summary').innerHTML, /3 tickets alimentan las métricas/);
assert.match(getElement('glpi-context-summary').innerHTML, /Grupo A/);
assert.match(getElement('glpi-context-summary').innerHTML, /Ana/);

api.toggleStatus('above_average');
assert.equal(api.state.status, 'above_average');
assert.deepEqual([...api.getFilteredTickets().map(item => item.id)], ['2']);
api.toggleStatus('above_average');
assert.equal(api.state.status, 'all');
assert.equal(api.getFilteredTickets().length, 3);

api.setStatus('closed');
assert.deepEqual([...api.getFilteredTickets().map(item => item.id)], ['3']);
api.setStatus('all');
api.renderFilterOptions();
assert.match(getElement('glpi-requester-filter').innerHTML, /Ana/);
assert.match(getElement('glpi-requester-filter').innerHTML, /Beto/);
assert.doesNotMatch(getElement('glpi-requester-filter').innerHTML, /Carla/);
assert.match(getElement('glpi-group-filter').innerHTML, /Grupo A/);
assert.doesNotMatch(getElement('glpi-group-filter').innerHTML, /Grupo B/);

api.setRequester('');
api.setStatus('above_average');
api.renderTable();
const filteredTable = getElement('glpi-ticket-table').innerHTML;
assert.match(filteredTable, /10 días/);
assert.match(filteredTable, /4,7 días sobre el promedio/);
assert.doesNotMatch(filteredTable, /<th>Días solución<\/th>/);
assert.doesNotMatch(filteredTable, /<th>Días abierto<\/th>/);

api.setStatus('all');
const summary = api.summarizeBy(
  api.state.tickets.filter(item => item.directorGroup === 'Grupo A'),
  item => item.directorGroup
)[0];
assert.equal(summary.pending, 3);
assert.equal(summary.averageOpenDays, 16 / 3);
assert.equal(summary.averageClosedDays, 4);

const closedTime = api.getTicketTimeInfo(api.state.tickets[2], api.getContextMetrics());
assert.equal(closedTime.days, 6);
assert.equal(closedTime.label, 'Días que tardó en cerrar');
assert.equal(closedTime.comparisonClass, 'over');

console.log('GLPI: cruces de grupo, solicitante, estado y tiempos validados.');
