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
  'renderDetail',
  'summarizeBy',
  'setGroup',
  'setRequester',
  'setStatus',
  'toggleStatus',
  'getTicketTimeInfo',
  'getStatusKey',
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
assert.equal(normalizedOpenTicket.statusKey, 'in_progress');

assert.equal(api.getStatusKey('Nuevo'), 'new');
assert.equal(api.getStatusKey('En curso (asignada)'), 'in_progress');
assert.equal(api.getStatusKey('En espera'), 'waiting');
assert.equal(api.getStatusKey('Resueltas'), 'resolved');
assert.equal(api.getStatusKey('Cerrado'), 'closed');

const normalizedClosedTicket = api.normalizeTicketRow({
  ID:9,
  'Fecha de Apertura':'2026-07-01',
  'Fecha de solución':'2026-07-09',
  'Fecha de cierre':'2026-07-12',
  Estado:'Cerrado',
  Solicitante:'Persona Prueba'
}, 1);
assert.equal(normalizedClosedTicket.daysSolution, 8, 'El cierre debe medir días entre apertura y solución.');
assert.equal(normalizedClosedTicket.daysClosed, 11, 'El cerrado debe medir días entre apertura y cierre.');

const closedWithoutDate = api.normalizeTicketRow({
  ID:10,
  'Fecha de Apertura':'2026-07-01',
  Estado:'Cerrado',
  Solicitante:'Persona Prueba'
}, 2);
assert.equal(closedWithoutDate.daysSolution, null, 'Un cierre sin fecha no debe inventar duración.');
assert.equal(closedWithoutDate.daysClosed, null, 'Un ticket cerrado sin fecha de cierre no debe inventar duración.');

function ticket(id, statusKey, group, requester, times = {}) {
  const statusLabels = {
    new:'Nuevo',
    in_progress:'En curso (asignada)',
    waiting:'En espera',
    resolved:'Resueltas',
    closed:'Cerrado'
  };
  return {
    id:String(id),
    title:`Ticket ${id}`,
    status:statusLabels[statusKey],
    statusKey,
    openDate:'2026-07-01',
    solutionDate:['resolved','closed'].includes(statusKey) ? '2026-07-10' : '',
    closeDate:statusKey === 'closed' ? '2026-07-12' : '',
    period:'2026-07',
    technicians:['Técnico Uno'],
    category:'Soporte > Prueba',
    categoryGroup:'Soporte',
    directorGroup:group,
    requester,
    description:'',
    daysOpen:times.daysOpen ?? null,
    daysSolution:times.daysSolution ?? null,
    daysClosed:times.daysClosed ?? times.daysSolution ?? null
  };
}

api.state.loaded = true;
api.state.period = '2026-07';
api.state.tickets = [
  ticket(1, 'new', 'Grupo A', 'Ana', { daysOpen:2 }),
  ticket(2, 'in_progress', 'Grupo A', 'Ana', { daysOpen:10 }),
  ticket(3, 'in_progress', 'Grupo A', 'Beto', { daysOpen:4 }),
  ticket(4, 'waiting', 'Grupo A', 'Beto', { daysOpen:8 }),
  ticket(5, 'resolved', 'Grupo A', 'Ana', { daysSolution:6 }),
  ticket(6, 'closed', 'Grupo A', 'Ana', { daysSolution:1, daysClosed:2 }),
  ticket(7, 'closed', 'Grupo A', 'Beto', { daysSolution:2, daysClosed:4 }),
  ticket(8, 'in_progress', 'Grupo B', 'Carla', { daysOpen:20 }),
  ticket(9, 'resolved', 'Grupo B', 'Carla', { daysSolution:10 }),
  ticket(10, 'closed', 'Grupo B', 'Carla', { daysSolution:8, daysClosed:12 })
];
api.state.tickets[0].description = 'Datos del formularioSection1) Empresa. : Pluxee Colombia SAS2) Labor a realizar: Primera línea\nSegunda línea\nObservación final del caso';

api.setGroup('Grupo A');
let metrics = api.getContextMetrics();
assert.equal(metrics.context.length, 7);
assert.equal(metrics.activeTickets.length, 4);
assert.equal(metrics.newTickets.length, 1);
assert.equal(metrics.inProgressTickets.length, 2);
assert.equal(metrics.waitingTickets.length, 1);
assert.equal(metrics.averageOpenDays, 6);
assert.equal(metrics.medianOpenDays, 6);
assert.equal(metrics.maxOpenDays, 10);
assert.equal(metrics.resolvedTickets.length, 1);
assert.equal(metrics.averageResolvedDays, 6);
assert.equal(metrics.closedTickets.length, 2);
assert.equal(metrics.averageClosedDays, 3);
assert.equal(metrics.medianClosedDays, 3);
assert.equal(metrics.maxClosedDays, 4);
assert.equal(metrics.statusMetrics.in_progress.averageDays, 7);
assert.equal(metrics.statusMetrics.waiting.averageDays, 8);

api.renderDetail(api.state.tickets[0]);
const detailMarkup = getElement('glpi-detail-content').innerHTML;
assert.match(detailMarkup, /Descripción completa del caso/);
assert.match(detailMarkup, /class="glpi-description-list"/);
assert.match(detailMarkup, /class="glpi-description-number">1<\/span>/);
assert.match(detailMarkup, /<strong>Empresa<\/strong>/);
assert.match(detailMarkup, /Segunda línea\nObservación final del caso/);
assert.match(detailMarkup, /Pluxee Colombia SAS/);
assert.doesNotMatch(detailMarkup, /Datos del formularioSection1\)/);
assert.doesNotMatch(detailMarkup, /Campos identificados/);

api.renderKpis();
const groupCards = getElement('glpi-kpis').innerHTML;
assert.match(groupCards, /Tickets del contexto[\s\S]*>7</);
assert.match(groupCards, /Tickets activos[\s\S]*>4</);
assert.match(groupCards, /Nuevo[\s\S]*>1</);
assert.match(groupCards, /En curso \(asignada\)[\s\S]*>2</);
assert.match(groupCards, /Prom\. 7 días abiertos/);
assert.match(groupCards, /En espera[\s\S]*>1</);
assert.match(groupCards, /Resueltas[\s\S]*>1</);
assert.match(groupCards, /Prom\. 6 días para resolverse/);
assert.match(groupCards, /Cerrado[\s\S]*>2</);
assert.match(groupCards, /Prom\. 3 días para cerrar/);
assert.doesNotMatch(groupCards, /Pendientes|sobre promedio/);

api.setRequester('Ana');
metrics = api.getContextMetrics();
assert.equal(metrics.context.length, 4);
assert.equal(metrics.activeTickets.length, 2);
assert.equal(metrics.averageOpenDays, 6);
assert.equal(metrics.averageResolvedDays, 6);
assert.equal(metrics.closedTickets.length, 1);
assert.equal(metrics.averageClosedDays, 2);
assert.equal(metrics.statusMetrics.in_progress.averageDays, 10);
assert.equal(metrics.statusMetrics.new.averageDays, 2);
assert.match(getElement('glpi-context-summary').innerHTML, /4 tickets alimentan las métricas/);
assert.match(getElement('glpi-context-summary').innerHTML, /Grupo A/);
assert.match(getElement('glpi-context-summary').innerHTML, /Ana/);

api.toggleStatus('in_progress');
assert.equal(api.state.status, 'in_progress');
assert.deepEqual([...api.getFilteredTickets().map(item => item.id)], ['2']);
api.toggleStatus('in_progress');
assert.equal(api.state.status, 'all');
assert.equal(api.getFilteredTickets().length, 4);

api.setStatus('closed');
assert.deepEqual([...api.getFilteredTickets().map(item => item.id)], ['6']);
api.setStatus('all');
api.renderFilterOptions();
assert.match(getElement('glpi-requester-filter').innerHTML, /Ana/);
assert.match(getElement('glpi-requester-filter').innerHTML, /Beto/);
assert.doesNotMatch(getElement('glpi-requester-filter').innerHTML, /Carla/);
assert.match(getElement('glpi-group-filter').innerHTML, /Grupo A/);
assert.doesNotMatch(getElement('glpi-group-filter').innerHTML, /Grupo B/);
const statusOptions = getElement('glpi-status-filter').innerHTML;
assert.match(statusOptions, /Nuevo/);
assert.match(statusOptions, /En curso \(asignada\)/);
assert.match(statusOptions, /En espera/);
assert.match(statusOptions, /Resueltas/);
assert.match(statusOptions, /Cerrado/);
assert.doesNotMatch(statusOptions, /Pendientes|sobre el promedio|Abiertos/);

api.setRequester('');
api.setStatus('in_progress');
api.renderTable();
const filteredTable = getElement('glpi-ticket-table').innerHTML;
assert.match(filteredTable, /10 días/);
assert.match(filteredTable, /3 días sobre el promedio/);
assert.doesNotMatch(filteredTable, /<th>Días solución<\/th>/);
assert.doesNotMatch(filteredTable, /<th>Días abierto<\/th>/);

api.setStatus('all');
const summary = api.summarizeBy(
  api.state.tickets.filter(item => item.directorGroup === 'Grupo A'),
  item => item.directorGroup
)[0];
assert.equal(summary.active, 4);
assert.equal(summary.averageOpenDays, 6);
assert.equal(summary.resolved, 1);
assert.equal(summary.averageResolvedDays, 6);
assert.equal(summary.closed, 2);
assert.equal(summary.averageClosedDays, 3);

const resolvedTime = api.getTicketTimeInfo(api.state.tickets[4], api.getContextMetrics());
assert.equal(resolvedTime.days, 6);
assert.equal(resolvedTime.label, 'Días que tardó en resolverse');

const closedTime = api.getTicketTimeInfo(api.state.tickets[5], api.getContextMetrics());
assert.equal(closedTime.days, 2);
assert.equal(closedTime.label, 'Días que tardó en cerrar');
assert.equal(closedTime.comparisonClass, 'under');

console.log('GLPI: cinco estados reales, cruces y tiempos por estado validados.');
