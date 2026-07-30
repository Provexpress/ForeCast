(function(){
  'use strict';

  const GLPI_FOLDER_NAME = 'GLPI';
  const GLPI_FILE_NAME = 'TICKETS-GLPI.xlsx';
  const GLPI_SHEET_NAME = '2026';
  const PAGE_SIZE = 60;
  let commercialPeopleCache = null;
  let statusChartInstance = null;

  const state = {
    loading: false,
    loaded: false,
    error: '',
    tickets: [],
    sourceRows: 0,
    duplicateRows: 0,
    sourceModifiedAt: '',
    period: '',
    status: 'all',
    group: '',
    requester: '',
    search: '',
    limit: PAGE_SIZE,
    selectedId: ''
  };

  function escapeHtml(value){
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeAttr(value){
    return escapeHtml(value).replace(/`/g, '&#96;');
  }

  function jsString(value){
    return JSON.stringify(String(value == null ? '' : value))
      .replace(/</g, '\\u003c')
      .replace(/>/g, '\\u003e')
      .replace(/&/g, '\\u0026');
  }

  function normalizeText(value){
    return String(value == null ? '' : value)
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
  }

  function cleanPlainText(value){
    return String(value == null ? '' : value)
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(?:p|div|section|article|li|tr|h[1-6])\s*>/gi, '\n')
      .replace(/<li(?:\s[^>]*)?>/gi, '• ')
      .replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/\r/g, '')
      .replace(/[ \t]+/g, ' ')
      .replace(/\n\s+/g, '\n')
      .trim();
  }

  function normalizeHeader(value){
    return normalizeText(value).replace(/[^a-z0-9]/g, '');
  }

  function buildRowLookup(row){
    const lookup = {};
    Object.entries(row || {}).forEach(([key, value]) => {
      lookup[normalizeHeader(key)] = value;
    });
    return lookup;
  }

  function getField(lookup, names){
    for(const name of names) {
      const value = lookup[normalizeHeader(name)];
      if(value !== null && value !== undefined && value !== '') return value;
    }
    return null;
  }

  function dateToIso(value){
    if(value instanceof Date && !Number.isNaN(value.getTime())) {
      return value.toISOString().slice(0, 10);
    }
    if(typeof value === 'number' && Number.isFinite(value)) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      epoch.setUTCDate(epoch.getUTCDate() + Math.floor(value));
      return epoch.toISOString().slice(0, 10);
    }
    const raw = cleanPlainText(value);
    if(!raw) return '';
    const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if(iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
    const dmy = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
    if(dmy) {
      const year = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3];
      return `${year}-${String(dmy[2]).padStart(2, '0')}-${String(dmy[1]).padStart(2, '0')}`;
    }
    const parsed = new Date(raw);
    return Number.isNaN(parsed.getTime()) ? '' : parsed.toISOString().slice(0, 10);
  }

  function normalizeTicketId(value){
    const raw = cleanPlainText(value);
    const digits = raw.replace(/\D/g, '');
    return digits || raw.replace(/\s+/g, '');
  }

  function parseNumber(value){
    if(value === null || value === undefined || value === '') return null;
    if(typeof value === 'number') return Number.isFinite(value) ? value : null;
    const parsed = Number(String(value).replace(',', '.').replace(/[^\d.-]/g, ''));
    return Number.isFinite(parsed) ? parsed : null;
  }

  function getBogotaDateKey(date){
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone:'America/Bogota',
      year:'numeric',
      month:'2-digit',
      day:'2-digit'
    }).formatToParts(date || new Date()).reduce((acc, part) => {
      if(part.type !== 'literal') acc[part.type] = part.value;
      return acc;
    }, {});
    return `${parts.year}-${parts.month}-${parts.day}`;
  }

  function diffDays(startIso, endIso){
    if(!startIso) return null;
    const start = new Date(`${startIso}T00:00:00Z`);
    const end = new Date(`${endIso || getBogotaDateKey()}T00:00:00Z`);
    if(Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return null;
    return Math.max(0, Math.round((end - start) / 86400000));
  }

  function getStatusKey(value){
    const normalized = normalizeText(value);
    if(normalized.includes('nuevo')) return 'new';
    if(normalized.includes('espera')) return 'waiting';
    if(normalized.includes('resuelt') || normalized.includes('solucion')) return 'resolved';
    if(normalized.includes('cerrad')) return 'closed';
    if(normalized.includes('curso') || normalized.includes('asignad')) return 'in_progress';
    return 'unknown';
  }

  function isActiveStatus(statusKey){
    return ['new','in_progress','waiting'].includes(statusKey);
  }

  function isTerminalStatus(statusKey){
    return ['resolved','closed'].includes(statusKey);
  }

  function splitPeople(value){
    const normalized = String(value == null ? '' : value)
      .replace(/<br\s*\/?>/gi, '|')
      .replace(/\n/g, '|');
    return [...new Set(normalized.split('|').map(cleanPersonName).filter(Boolean))];
  }

  function cleanPersonName(value){
    return cleanPlainText(value).replace(/\s*\(\d+\)\s*$/, '').trim();
  }

  function getCategoryGroup(value){
    const category = cleanPlainText(value) || 'Sin categoría';
    return category.split('>')[0].trim() || 'Sin categoría';
  }

  function getPersonTokens(value){
    return normalizeText(value)
      .replace(/[^a-z0-9\s]/g, ' ')
      .split(/\s+/)
      .filter(Boolean);
  }

  function getCommercialPeople(){
    if(commercialPeopleCache) return commercialPeopleCache;
    const structure = window.FORECAST_STRUCTURE || {};
    const source = structure.estructura || {};
    const people = [];
    const addEntries = (entries, role) => {
      Object.entries(entries || {}).forEach(([email, data]) => {
        const groupNumber = Number(data.grupo || (structure.getGroupByEmail && structure.getGroupByEmail(email)));
        if(!groupNumber) return;
        const directorName = structure.getDirectorNameByGroup
          ? structure.getDirectorNameByGroup(groupNumber)
          : '';
        const aliases = [
          data.nombre,
          String(data.archivo || '').replace(/\.(xlsx|xls)$/i, '')
        ].filter(Boolean);
        people.push({
          email,
          role,
          groupNumber,
          directorName: directorName || `Director grupo ${groupNumber}`,
          aliases
        });
      });
    };
    addEntries(source.directores, 'director');
    addEntries(source.ejecutivos, 'ejecutivo');
    addEntries(source.salesSupport, 'sales_support');
    addEntries(source.salesSupportComerciales, 'sales_support_comercial');
    commercialPeopleCache = people;
    return people;
  }

  function scorePersonAlias(requester, alias){
    const requesterNormalized = normalizeText(requester).replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
    const aliasNormalized = normalizeText(alias).replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
    if(!requesterNormalized || !aliasNormalized) return 0;
    if(requesterNormalized === aliasNormalized) return 1000;
    const requesterTokens = new Set(getPersonTokens(requester));
    const aliasTokens = [...new Set(getPersonTokens(alias))];
    if(aliasTokens.length < 2) return 0;
    const common = aliasTokens.filter(token => requesterTokens.has(token));
    if(common.length < 2) return 0;
    const coverage = common.length / aliasTokens.length;
    if(coverage < .66) return 0;
    const contained = aliasTokens.every(token => requesterTokens.has(token));
    return (contained ? 700 : 400) + common.length * 20 + coverage * 10 - Math.max(0, requesterTokens.size - aliasTokens.length);
  }

  function resolveDirectorGroup(requester){
    let best = null;
    getCommercialPeople().forEach(person => {
      const score = Math.max(0, ...person.aliases.map(alias => scorePersonAlias(requester, alias)));
      if(score && (!best || score > best.score)) best = { person, score };
    });
    if(!best) {
      return {
        directorGroup: 'Sin grupo comercial',
        directorName: '',
        matchedPerson: ''
      };
    }
    return {
      directorGroup: `Grupo ${best.person.directorName}`,
      directorName: best.person.directorName,
      matchedPerson: best.person.aliases[0] || ''
    };
  }

  function normalizeTicketRow(row, index){
    const lookup = buildRowLookup(row);
    const id = normalizeTicketId(getField(lookup, ['ID']));
    const openDate = dateToIso(getField(lookup, ['Fecha de Apertura']));
    const solutionDate = dateToIso(getField(lookup, ['Fecha de solución', 'Fecha de solucion']));
    const closeDate = dateToIso(getField(lookup, ['Fecha de cierre', 'Fecha de cerrado']));
    const status = cleanPlainText(getField(lookup, ['Estado'])) || 'Sin estado';
    const statusKey = getStatusKey(status);
    const requester = cleanPersonName(getField(lookup, ['Solicitante - Solicitante', 'Solicitante'])) || 'Sin solicitante';
    const category = cleanPlainText(getField(lookup, ['Categoría', 'Categoria'])) || 'Sin categoría';
    const directorGroup = resolveDirectorGroup(requester);
    const technicians = splitPeople(getField(lookup, ['Asignado a: - Técnico', 'Asignado a - Técnico', 'Técnico']));
    const sourceDaysSolution = parseNumber(getField(lookup, ['Días en Solución', 'Dias en Solucion']));
    const sourceDaysClosed = parseNumber(getField(lookup, ['Días en Cierre', 'Dias en Cierre', 'Días para cerrar', 'Dias para cerrar']));
    const sourceDaysOpen = parseNumber(getField(lookup, ['Días Abierto', 'Dias Abierto']));
    const daysSolution = sourceDaysSolution === null
      ? (solutionDate ? diffDays(openDate, solutionDate) : null)
      : sourceDaysSolution;
    const daysClosed = sourceDaysClosed === null
      ? (closeDate ? diffDays(openDate, closeDate) : daysSolution)
      : sourceDaysClosed;
    const calculatedDaysOpen = diffDays(openDate);
    const daysOpen = isTerminalStatus(statusKey)
      ? null
      : (calculatedDaysOpen === null
        ? sourceDaysOpen
        : Math.max(calculatedDaysOpen, sourceDaysOpen === null ? 0 : sourceDaysOpen));

    return {
      id: id || `fila-${index + 1}`,
      title: cleanPlainText(getField(lookup, ['Título', 'Titulo'])) || 'Ticket sin título',
      status,
      statusKey,
      openDate,
      solutionDate,
      closeDate,
      period: openDate.slice(0, 7),
      technicians,
      category,
      categoryGroup: getCategoryGroup(category),
      directorGroup: directorGroup.directorGroup,
      directorName: directorGroup.directorName,
      matchedCommercialPerson: directorGroup.matchedPerson,
      description: cleanPlainText(getField(lookup, ['Descripción', 'Descripcion'])),
      requester,
      daysSolution,
      daysClosed,
      daysOpen
    };
  }

  function mergeTickets(current, candidate){
    const merged = { ...current };
    const preferCandidate = [
      candidate.title,
      candidate.status,
      candidate.openDate,
      candidate.solutionDate,
      candidate.closeDate,
      candidate.category,
      candidate.description,
      candidate.requester
    ].filter(Boolean).length > [
      current.title,
      current.status,
      current.openDate,
      current.solutionDate,
      current.closeDate,
      current.category,
      current.description,
      current.requester
    ].filter(Boolean).length;
    const preferred = preferCandidate ? candidate : current;
    Object.assign(merged, preferred);
    merged.technicians = [...new Set([...(current.technicians || []), ...(candidate.technicians || [])])];
    if((candidate.description || '').length > (current.description || '').length) merged.description = candidate.description;
    if(current.daysSolution !== null && current.daysSolution !== undefined) merged.daysSolution = current.daysSolution;
    if(current.daysClosed !== null && current.daysClosed !== undefined) merged.daysClosed = current.daysClosed;
    if(current.daysOpen !== null && current.daysOpen !== undefined) merged.daysOpen = current.daysOpen;
    return merged;
  }

  function deduplicateTickets(rows){
    const byId = new Map();
    rows.forEach((row, index) => {
      const ticket = normalizeTicketRow(row, index);
      if(!ticket.openDate && !ticket.title && !ticket.id) return;
      if(byId.has(ticket.id)) byId.set(ticket.id, mergeTickets(byId.get(ticket.id), ticket));
      else byId.set(ticket.id, ticket);
    });
    return [...byId.values()].sort((a, b) =>
      String(b.openDate).localeCompare(String(a.openDate)) || String(b.id).localeCompare(String(a.id), undefined, { numeric:true })
    );
  }

  function getCurrentPeriod(){
    return getBogotaDateKey().slice(0, 7);
  }

  function getPeriods(){
    return [...new Set([
      getCurrentPeriod(),
      ...state.tickets.map(ticket => ticket.period)
    ].filter(period => /^\d{4}-\d{2}$/.test(String(period || ''))))].sort().reverse();
  }

  function ensureDefaultPeriod(){
    const periods = getPeriods();
    if(state.period && (state.period === 'all' || periods.includes(state.period))) return;
    state.period = getCurrentPeriod();
  }

  function formatMonth(period){
    if(!period || period === 'all') return 'Todos los meses (Visión Global)';
    const [year, month] = period.split('-').map(Number);
    if(!year || !month) return String(period);
    const date = new Date(Date.UTC(year, month - 1, 1));
    return new Intl.DateTimeFormat('es-CO', { month:'long', year:'numeric', timeZone:'UTC' }).format(date);
  }

  function formatDate(iso){
    if(!iso) return '—';
    const date = new Date(`${iso}T00:00:00Z`);
    return new Intl.DateTimeFormat('es-CO', { day:'2-digit', month:'short', year:'numeric', timeZone:'UTC' }).format(date);
  }

  function formatSourceDate(value){
    if(!value) return 'sin fecha de actualización';
    const date = new Date(value);
    if(Number.isNaN(date.getTime())) return cleanPlainText(value);
    return new Intl.DateTimeFormat('es-CO', {
      dateStyle:'medium',
      timeStyle:'short',
      timeZone:'America/Bogota'
    }).format(date);
  }

  function formatDays(value){
    if(value === null || value === undefined || !Number.isFinite(Number(value))) return '—';
    return Number(value).toLocaleString('es-CO', { maximumFractionDigits:1 });
  }

  function formatDaysWithUnit(value){
    return isFiniteMetric(value) ? `${formatDays(value)} días` : 'Sin dato';
  }

  function getFiniteNumbers(values){
    return (values || [])
      .filter(value => value !== null && value !== undefined && value !== '')
      .map(Number)
      .filter(Number.isFinite);
  }

  function isFiniteMetric(value){
    return value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value));
  }

  function average(values){
    const valid = getFiniteNumbers(values);
    return valid.length ? valid.reduce((sum, value) => sum + value, 0) / valid.length : null;
  }

  function median(values){
    const valid = getFiniteNumbers(values).sort((a, b) => a - b);
    if(!valid.length) return null;
    const middle = Math.floor(valid.length / 2);
    return valid.length % 2 ? valid[middle] : (valid[middle - 1] + valid[middle]) / 2;
  }

  function maximum(values){
    const valid = getFiniteNumbers(values);
    return valid.length ? Math.max(...valid) : null;
  }

  function getMonthTickets(){
    return state.tickets.filter(ticket => !state.period || state.period === 'all' || ticket.period === state.period);
  }

  function matchesSearch(ticket){
    if(!state.search) return true;
    const haystack = normalizeText([
      ticket.id,
      ticket.title,
      ticket.status,
      ticket.category,
      ticket.description,
      ticket.requester,
      ticket.technicians.join(' ')
    ].join(' '));
    return haystack.includes(normalizeText(state.search));
  }

  function getContextTickets(){
    return getMonthTickets().filter(ticket =>
      (!state.group || ticket.directorGroup === state.group) &&
      (!state.requester || ticket.requester === state.requester) &&
      matchesSearch(ticket)
    );
  }

  function getContextMetrics(tickets){
    const context = tickets || getContextTickets();
    const newTickets = context.filter(ticket => ticket.statusKey === 'new');
    const inProgressTickets = context.filter(ticket => ticket.statusKey === 'in_progress');
    const waitingTickets = context.filter(ticket => ticket.statusKey === 'waiting');
    const activeTickets = context.filter(ticket => isActiveStatus(ticket.statusKey));
    const resolvedTickets = context.filter(ticket => ticket.statusKey === 'resolved');
    const closedTickets = context.filter(ticket => ticket.statusKey === 'closed');
    const activeDays = getFiniteNumbers(activeTickets.map(ticket => ticket.daysOpen));
    const resolvedDays = getFiniteNumbers(resolvedTickets.map(ticket => ticket.daysSolution));
    const closedDays = getFiniteNumbers(closedTickets.map(ticket => ticket.daysClosed));
    const averageOpenDays = average(activeDays);
    const averageResolvedDays = average(resolvedDays);
    const averageClosedDays = average(closedDays);
    const buildStatusMetrics = (statusTickets, field) => {
      const values = getFiniteNumbers(statusTickets.map(ticket => ticket[field]));
      return {
        count:statusTickets.length,
        averageDays:average(values),
        medianDays:median(values),
        maxDays:maximum(values)
      };
    };
    return {
      context,
      newTickets,
      inProgressTickets,
      waitingTickets,
      activeTickets,
      resolvedTickets,
      closedTickets,
      averageOpenDays,
      medianOpenDays:median(activeDays),
      maxOpenDays:maximum(activeDays),
      averageResolvedDays,
      medianResolvedDays:median(resolvedDays),
      maxResolvedDays:maximum(resolvedDays),
      averageClosedDays,
      medianClosedDays:median(closedDays),
      maxClosedDays:maximum(closedDays),
      statusMetrics:{
        new:buildStatusMetrics(newTickets, 'daysOpen'),
        in_progress:buildStatusMetrics(inProgressTickets, 'daysOpen'),
        waiting:buildStatusMetrics(waitingTickets, 'daysOpen'),
        resolved:buildStatusMetrics(resolvedTickets, 'daysSolution'),
        closed:buildStatusMetrics(closedTickets, 'daysClosed')
      }
    };
  }

  function filterTicketsByStatus(tickets, status){
    const context = tickets || [];
    const selected = status || state.status;
    if(selected === 'all') return context;
    if(selected === 'active') return context.filter(ticket => isActiveStatus(ticket.statusKey));
    if(selected === 'critical') return context.filter(ticket => isActiveStatus(ticket.statusKey) && (ticket.daysOpen || 0) > 15);
    return context.filter(ticket => ticket.statusKey === selected);
  }

  function getFilteredTickets(){
    return filterTicketsByStatus(getContextTickets(), state.status);
  }

  function getStatusLabel(key){
    return ({
      all:'Todos los estados',
      active:'Tickets activos',
      critical:'🔥 Casos Críticos (>15d)',
      new:'Nuevo',
      in_progress:'En curso (asignada)',
      waiting:'En espera',
      resolved:'Resueltas',
      closed:'Cerrado'
    })[key] || 'Sin estado';
  }

  function toggleStatus(value){
    const valid = ['all','active','critical','new','in_progress','waiting','resolved','closed'];
    const selected = valid.includes(value) ? value : 'all';
    state.status = selected !== 'all' && state.status === selected ? 'all' : selected;
    resetPaging();
    renderAll();
  }

  function getStatusClass(key){
    return ({
      new:'new',
      in_progress:'in-progress',
      waiting:'waiting',
      resolved:'resolved',
      closed:'closed'
    })[key] || 'unknown';
  }

  function renderLoading(){
    const kpis = document.getElementById('glpi-kpis');
    const table = document.getElementById('glpi-ticket-table');
    const source = document.getElementById('glpi-source-note');
    const context = document.getElementById('glpi-context-summary');
    if(source) source.textContent = 'Conectando con TICKETS-GLPI.xlsx en SharePoint…';
    if(context) {
      context.innerHTML = '';
      context.hidden = true;
    }
    if(kpis) kpis.innerHTML = '<div class="glpi-loading-card">Leyendo y organizando tickets…</div>';
    if(table) table.innerHTML = '<div class="glpi-empty-state">Preparando datos del Excel.</div>';
  }

  function renderError(){
    const message = state.error || 'No fue posible cargar los tickets.';
    const source = document.getElementById('glpi-source-note');
    const kpis = document.getElementById('glpi-kpis');
    const table = document.getElementById('glpi-ticket-table');
    const context = document.getElementById('glpi-context-summary');
    if(source) source.textContent = message;
    if(context) {
      context.innerHTML = '';
      context.hidden = true;
    }
    if(kpis) kpis.innerHTML = `<div class="glpi-error-card">${escapeHtml(message)}</div>`;
    if(table) table.innerHTML = '<div class="glpi-empty-state">Usa “Actualizar Excel” para intentar de nuevo.</div>';
  }

  function renderPeriodOptions(){
    const select = document.getElementById('glpi-period');
    if(!select) return;
    const periods = getPeriods();
    const options = [
      `<option value="all"${state.period === 'all' || !state.period ? ' selected' : ''}>Todos los meses (Visión Global)</option>`,
      ...periods.map(period => `<option value="${escapeAttr(period)}"${period === state.period ? ' selected' : ''}>${escapeHtml(formatMonth(period))}</option>`)
    ];
    select.innerHTML = options.join('');
  }

  function renderFilterOptions(){
    const monthTickets = getMonthTickets();
    const allGroups = [...new Set(monthTickets.map(ticket => ticket.directorGroup).filter(Boolean))];
    const allRequesters = [...new Set(monthTickets.map(ticket => ticket.requester).filter(Boolean))];
    const groupSelect = document.getElementById('glpi-group-filter');
    const requesterSelect = document.getElementById('glpi-requester-filter');
    const statusSelect = document.getElementById('glpi-status-filter');
    const searchInput = document.getElementById('glpi-search');

    if(state.group && !allGroups.includes(state.group)) state.group = '';
    if(state.requester && !allRequesters.includes(state.requester)) state.requester = '';
    const searchTickets = monthTickets.filter(matchesSearch);
    const groups = [...new Set(searchTickets
      .filter(ticket => !state.requester || ticket.requester === state.requester)
      .map(ticket => ticket.directorGroup)
      .filter(Boolean))];
    const requesters = [...new Set(searchTickets
      .filter(ticket => !state.group || ticket.directorGroup === state.group)
      .map(ticket => ticket.requester)
      .filter(Boolean))];
    if(state.group && !groups.includes(state.group)) groups.push(state.group);
    if(state.requester && !requesters.includes(state.requester)) requesters.push(state.requester);
    groups.sort((a, b) => a.localeCompare(b, 'es'));
    requesters.sort((a, b) => a.localeCompare(b, 'es'));

    if(groupSelect) {
      groupSelect.innerHTML = '<option value="">Todos los grupos de directores</option>' + groups
        .map(group => `<option value="${escapeAttr(group)}"${group === state.group ? ' selected' : ''}>${escapeHtml(group)}</option>`)
        .join('');
      groupSelect.value = state.group;
    }
    if(requesterSelect) {
      requesterSelect.innerHTML = '<option value="">Todos los solicitantes</option>' + requesters
        .map(requester => `<option value="${escapeAttr(requester)}"${requester === state.requester ? ' selected' : ''}>${escapeHtml(requester)}</option>`)
        .join('');
      requesterSelect.value = state.requester;
    }
    if(statusSelect) {
      statusSelect.innerHTML = [
        ['all','Todos los estados'],
        ['new','Nuevo'],
        ['in_progress','En curso (asignada)'],
        ['waiting','En espera'],
        ['resolved','Resueltas'],
        ['closed','Cerrado']
      ].map(([value, label]) => `<option value="${value}"${value === state.status ? ' selected' : ''}>${label}</option>`).join('');
      statusSelect.value = state.status;
    }
    if(searchInput && searchInput.value !== state.search) searchInput.value = state.search;
  }

  function renderContextSummary(){
    const host = document.getElementById('glpi-context-summary');
    if(!host) return;
    host.hidden = false;
    const context = getContextTickets();
    const chips = [
      `<span class="glpi-context-chip glpi-context-period">${escapeHtml(formatMonth(state.period))}</span>`,
      state.group
        ? `<button type="button" class="glpi-context-chip" onclick="GlpiModule.setGroup('')" title="Quitar grupo">${escapeHtml(state.group)} <strong>×</strong></button>`
        : '<span class="glpi-context-chip">Todos los grupos</span>',
      state.requester
        ? `<button type="button" class="glpi-context-chip" onclick="GlpiModule.setRequester('')" title="Quitar solicitante">${escapeHtml(state.requester)} <strong>×</strong></button>`
        : '<span class="glpi-context-chip">Todos los solicitantes</span>',
      state.status !== 'all'
        ? `<button type="button" class="glpi-context-chip" onclick="GlpiModule.setStatus('all')" title="Quitar estado">${escapeHtml(getStatusLabel(state.status))} <strong>×</strong></button>`
        : '',
      state.search
        ? `<button type="button" class="glpi-context-chip" onclick="GlpiModule.setSearch('')" title="Quitar búsqueda">Búsqueda: ${escapeHtml(state.search)} <strong>×</strong></button>`
        : ''
    ].filter(Boolean).join('');
    host.innerHTML = `
      <div>
        <span class="glpi-context-eyebrow">Contexto de análisis</span>
        <strong>${context.length.toLocaleString('es-CO')} tickets alimentan las métricas</strong>
      </div>
      <div class="glpi-context-chips">${chips}</div>`;
  }

  function kpiCard(kind, label, value, detail, filterValue){
    const clickable = Boolean(filterValue);
    const active = clickable && state.status === filterValue;
    const tag = clickable ? 'button' : 'article';
    const action = clickable
      ? ` type="button" onclick="GlpiModule.toggleStatus('${escapeAttr(filterValue)}')" aria-pressed="${active ? 'true' : 'false'}"`
      : '';
    return `<${tag} class="glpi-kpi-card glpi-kpi-${escapeAttr(kind)}${active ? ' active' : ''}"${action}>
      <span class="glpi-kpi-label">${escapeHtml(label)}</span>
      <span class="glpi-kpi-value-line"><strong>${escapeHtml(value)}</strong></span>
      <small>${escapeHtml(detail)}</small>
    </${tag}>`;
  }

  function formatStatusTimeDetail(metrics, action){
    if(!metrics.count) return 'Sin tickets en este estado';
    if(metrics.averageDays === null) return `Sin tiempo calculable · ${metrics.count} tickets`;
    return `Prom. ${formatDaysWithUnit(metrics.averageDays)} ${action} · máximo ${formatDaysWithUnit(metrics.maxDays)}`;
  }

  function renderStatusDonutChart(metrics){
    const canvas = document.getElementById('glpi-status-donut-chart');
    const centerEl = document.getElementById('glpi-donut-center');
    const statsEl = document.getElementById('glpi-donut-stats');
    if(!canvas) return;

    if(statusChartInstance) {
      statusChartInstance.destroy();
      statusChartInstance = null;
    }

    const total = metrics.context.length || 0;
    const resolved = metrics.resolvedTickets.length;
    const closed = metrics.closedTickets.length;
    const resolvedTotal = resolved + closed;
    const inProgress = metrics.inProgressTickets.length;
    const newCount = metrics.newTickets.length;
    const waiting = metrics.waitingTickets.length;
    const criticalCount = metrics.activeTickets.filter(t => (t.daysOpen || 0) > 15).length;
    const pctResolved = total ? (resolvedTotal / total * 100).toFixed(1) : '0';

    if(centerEl) {
      centerEl.innerHTML = `
        <span class="glpi-center-val">${pctResolved}%</span>
        <span class="glpi-center-lbl">Resueltos</span>
      `;
    }

    if(statsEl) {
      const activeStatus = state.status;
      statsEl.innerHTML = `
        <button type="button" class="glpi-stat-pill glpi-stat-success${activeStatus === 'resolved' || activeStatus === 'closed' ? ' active' : ''}" onclick="GlpiModule.toggleStatus('resolved')" aria-pressed="${activeStatus === 'resolved'}">
          <span class="glpi-stat-dot"></span>
          <span class="glpi-stat-text">Resueltos / Cerrados: <strong>${resolvedTotal.toLocaleString('es-CO')}</strong> (${pctResolved}%)</span>
        </button>
        <button type="button" class="glpi-stat-pill glpi-stat-warning${activeStatus === 'in_progress' ? ' active' : ''}" onclick="GlpiModule.toggleStatus('in_progress')" aria-pressed="${activeStatus === 'in_progress'}">
          <span class="glpi-stat-dot"></span>
          <span class="glpi-stat-text">En Curso (Asignada): <strong>${inProgress.toLocaleString('es-CO')}</strong></span>
        </button>
        <button type="button" class="glpi-stat-pill glpi-stat-info${activeStatus === 'new' ? ' active' : ''}" onclick="GlpiModule.toggleStatus('new')" aria-pressed="${activeStatus === 'new'}">
          <span class="glpi-stat-dot"></span>
          <span class="glpi-stat-text">Nuevos: <strong>${newCount.toLocaleString('es-CO')}</strong></span>
        </button>
        <button type="button" class="glpi-stat-pill glpi-stat-purple${activeStatus === 'waiting' ? ' active' : ''}" onclick="GlpiModule.toggleStatus('waiting')" aria-pressed="${activeStatus === 'waiting'}">
          <span class="glpi-stat-dot"></span>
          <span class="glpi-stat-text">En Espera: <strong>${waiting.toLocaleString('es-CO')}</strong></span>
        </button>
        ${criticalCount > 0 ? `
        <button type="button" class="glpi-stat-pill glpi-stat-danger${activeStatus === 'critical' ? ' active' : ''}" onclick="GlpiModule.toggleStatus('critical')" aria-pressed="${activeStatus === 'critical'}">
          <span class="glpi-stat-dot"></span>
          <span class="glpi-stat-text">🔥 Críticos (&gt;15 días): <strong>${criticalCount.toLocaleString('es-CO')}</strong></span>
        </button>` : ''}
      `;
    }

    if(typeof Chart === 'undefined') return;

    const isLight = document.documentElement.getAttribute('data-theme') === 'light';

    statusChartInstance = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: ['Resueltos / Cerrados', 'En curso', 'Nuevos', 'En espera'],
        datasets: [{
          data: [resolvedTotal, inProgress, newCount, waiting],
          backgroundColor: [
            '#0DBF82',
            '#F0A020',
            '#2ABFDF',
            '#A77BDD'
          ],
          borderWidth: 3,
          borderColor: isLight ? '#FFFFFF' : '#111731',
          hoverOffset: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '72%',
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function(context) {
                const val = context.raw || 0;
                const pct = total ? (val / total * 100).toFixed(1) : 0;
                return ` ${context.label}: ${val} (${pct}%)`;
              }
            }
          }
        }
      }
    });
  }

  function renderKpis(){
    const host = document.getElementById('glpi-kpis');
    if(!host) return;
    const metrics = getContextMetrics();
    const activeDetail = metrics.averageOpenDays === null
      ? 'Sin antigüedad calculable'
      : `Prom. ${formatDaysWithUnit(metrics.averageOpenDays)} abiertos · máximo ${formatDaysWithUnit(metrics.maxOpenDays)}`;
    const criticalTickets = metrics.activeTickets.filter(t => (t.daysOpen || 0) > 15);
    const criticalDetail = criticalTickets.length ? `${criticalTickets.length} con >15 días sin solución` : 'Sin casos críticos de alta antigüedad';

    host.innerHTML = [
      kpiCard('total', 'Tickets del contexto', metrics.context.length.toLocaleString('es-CO'), formatMonth(state.period), 'all'),
      kpiCard('active', 'Tickets activos', metrics.activeTickets.length.toLocaleString('es-CO'), activeDetail, 'active'),
      kpiCard('critical', '🔥 Casos Críticos (>15d)', criticalTickets.length.toLocaleString('es-CO'), criticalDetail, 'critical'),
      kpiCard('new', 'Nuevo', metrics.newTickets.length.toLocaleString('es-CO'), formatStatusTimeDetail(metrics.statusMetrics.new, 'abiertos'), 'new'),
      kpiCard('in-progress', 'En curso (asignada)', metrics.inProgressTickets.length.toLocaleString('es-CO'), formatStatusTimeDetail(metrics.statusMetrics.in_progress, 'abiertos'), 'in_progress'),
      kpiCard('waiting', 'En espera', metrics.waitingTickets.length.toLocaleString('es-CO'), formatStatusTimeDetail(metrics.statusMetrics.waiting, 'abiertos'), 'waiting'),
      kpiCard('resolved', 'Resueltas', metrics.resolvedTickets.length.toLocaleString('es-CO'), formatStatusTimeDetail(metrics.statusMetrics.resolved, 'para resolverse'), 'resolved'),
      kpiCard('closed', 'Cerrado', metrics.closedTickets.length.toLocaleString('es-CO'), formatStatusTimeDetail(metrics.statusMetrics.closed, 'para cerrar'), 'closed')
    ].join('');
  }

  function summarizeBy(tickets, getter){
    const groups = new Map();
    tickets.forEach(ticket => {
      const name = getter(ticket) || 'Sin grupo';
      if(!groups.has(name)) groups.set(name, {
        name,
        total:0,
        new:0,
        in_progress:0,
        waiting:0,
        resolved:0,
        closed:0,
        openDays:[],
        resolvedDays:[],
        closedDays:[]
      });
      const group = groups.get(name);
      group.total += 1;
      if(Object.prototype.hasOwnProperty.call(group, ticket.statusKey)) group[ticket.statusKey] += 1;
      if(isActiveStatus(ticket.statusKey) && isFiniteMetric(ticket.daysOpen)) group.openDays.push(Number(ticket.daysOpen));
      if(ticket.statusKey === 'resolved' && isFiniteMetric(ticket.daysSolution)) group.resolvedDays.push(Number(ticket.daysSolution));
      if(ticket.statusKey === 'closed' && isFiniteMetric(ticket.daysClosed)) group.closedDays.push(Number(ticket.daysClosed));
    });
    return [...groups.values()]
      .map(group => ({
        ...group,
        active:group.new + group.in_progress + group.waiting,
        averageOpenDays:average(group.openDays),
        averageResolvedDays:average(group.resolvedDays),
        averageClosedDays:average(group.closedDays)
      }))
      .sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, 'es'));
  }

  function renderRankedList(hostId, groups, filterType){
    const host = document.getElementById(hostId);
    if(!host) return;
    if(!groups.length) {
      host.innerHTML = '<div class="glpi-empty-state">Sin datos para los filtros seleccionados.</div>';
      return;
    }
    const max = Math.max(...groups.map(group => group.total), 1);
    host.innerHTML = groups.slice(0, 12).map(group => {
      const active = filterType === 'group' ? state.group === group.name : state.requester === group.name;
      const action = filterType === 'group'
        ? `GlpiModule.toggleGroup(${jsString(group.name)})`
        : `GlpiModule.toggleRequester(${jsString(group.name)})`;
      return `<button class="glpi-rank-row${active ? ' active' : ''}" type="button" onclick="${escapeAttr(action)}" aria-pressed="${active ? 'true' : 'false'}">
        <span class="glpi-rank-main">
          <span class="glpi-rank-name" title="${escapeAttr(group.name)}">${escapeHtml(group.name)}</span>
          <span class="glpi-rank-value">${group.total.toLocaleString('es-CO')}</span>
        </span>
        <span class="glpi-rank-track"><span style="width:${Math.max(4, group.total / max * 100).toFixed(1)}%"></span></span>
        <span class="glpi-rank-meta">
          <span><strong>${group.active}</strong> activos · ${formatDaysWithUnit(group.averageOpenDays)} abiertos</span>
          <span><strong>${group.resolved}</strong> resueltas · ${formatDaysWithUnit(group.averageResolvedDays)} para resolverse</span>
          <span><strong>${group.closed}</strong> cerrados · ${formatDaysWithUnit(group.averageClosedDays)} para cerrar</span>
        </span>
      </button>`;
    }).join('');
  }

  function getDimensionTickets(filterType){
    let tickets = getMonthTickets().filter(matchesSearch);
    if(filterType !== 'group' && state.group) tickets = tickets.filter(ticket => ticket.directorGroup === state.group);
    if(filterType !== 'requester' && state.requester) tickets = tickets.filter(ticket => ticket.requester === state.requester);
    return filterTicketsByStatus(tickets, state.status);
  }

  function renderGroups(){
    renderRankedList(
      'glpi-category-groups',
      summarizeBy(getDimensionTickets('group'), ticket => ticket.directorGroup),
      'group'
    );
    renderRankedList(
      'glpi-requester-groups',
      summarizeBy(getDimensionTickets('requester'), ticket => ticket.requester),
      'requester'
    );
  }

  function getTicketTimeInfo(ticket, metrics){
    const days = ticket.statusKey === 'closed'
      ? ticket.daysClosed
      : (ticket.statusKey === 'resolved' ? ticket.daysSolution : ticket.daysOpen);
    const statusMetric = metrics.statusMetrics[ticket.statusKey];
    const averageDays = statusMetric ? statusMetric.averageDays : null;
    const delta = isFiniteMetric(days) && isFiniteMetric(averageDays)
      ? Number(days) - Number(averageDays)
      : null;
    const label = ticket.statusKey === 'resolved'
      ? 'Días que tardó en resolverse'
      : (ticket.statusKey === 'closed' ? 'Días que tardó en cerrar' : 'Días que lleva abierto');
    return {
      days,
      averageDays,
      delta,
      label,
      comparisonClass:delta === null || Math.abs(delta) < .05 ? 'neutral' : (delta > 0 ? 'over' : 'under'),
      comparison:delta === null
        ? 'Sin promedio comparable'
        : (Math.abs(delta) < .05
          ? 'En el promedio'
          : `${formatDays(Math.abs(delta))} días ${delta > 0 ? 'sobre' : 'bajo'} el promedio`)
    };
  }

  function renderTicketTimeCell(ticket, metrics){
    const time = getTicketTimeInfo(ticket, metrics);
    return `<span class="glpi-time-cell">
      <strong>${escapeHtml(formatDaysWithUnit(time.days))}</strong>
      <span>${escapeHtml(time.label)}</span>
      <em class="glpi-time-${time.comparisonClass}">${escapeHtml(time.comparison)}</em>
    </span>`;
  }

  function renderTable(){
    const host = document.getElementById('glpi-ticket-table');
    const meta = document.getElementById('glpi-table-meta');
    const more = document.getElementById('glpi-table-more');
    if(!host) return;
    const tickets = getFilteredTickets();
    const metrics = getContextMetrics();
    const visible = tickets.slice(0, state.limit);
    if(meta) meta.textContent = `${tickets.length.toLocaleString('es-CO')} tickets · ${getStatusLabel(state.status)} · clic en una fila para ver el detalle`;
    if(!visible.length) {
      host.innerHTML = '<div class="glpi-empty-state">No hay tickets para los filtros seleccionados.</div>';
      if(more) more.innerHTML = '';
      return;
    }

    host.innerHTML = `<table class="responsive-table glpi-table">
      <thead><tr>
        <th>ID</th><th>Título</th><th>Estado</th><th>Apertura</th><th>Grupo director</th>
        <th>Categoría</th><th>Solicitante</th><th>Técnico</th><th>Tiempo del ticket</th>
      </tr></thead>
      <tbody>${visible.map(ticket => `<tr class="glpi-ticket-row" role="button" onclick="GlpiModule.openTicket('${escapeAttr(ticket.id)}')" tabindex="0" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();GlpiModule.openTicket('${escapeAttr(ticket.id)}')}">
        <td class="td-mono" data-label="ID">#${escapeHtml(ticket.id)}</td>
        <td class="glpi-title-cell" data-label="Título" title="${escapeAttr(ticket.title)}">${escapeHtml(ticket.title)}</td>
        <td data-label="Estado"><span class="glpi-status glpi-status-${getStatusClass(ticket.statusKey)}">${escapeHtml(ticket.status)}</span></td>
        <td class="td-mono" data-label="Apertura">${escapeHtml(formatDate(ticket.openDate))}</td>
        <td data-label="Grupo director">${escapeHtml(ticket.directorGroup)}</td>
        <td data-label="Categoría">${escapeHtml(ticket.categoryGroup)}</td>
        <td data-label="Solicitante">${escapeHtml(ticket.requester)}</td>
        <td data-label="Técnico">${escapeHtml(ticket.technicians.join(', ') || 'Sin asignar')}</td>
        <td data-label="Tiempo del ticket">${renderTicketTimeCell(ticket, metrics)}</td>
      </tr>`).join('')}</tbody>
    </table>`;

    if(more) {
      more.innerHTML = tickets.length > visible.length
        ? `<button type="button" onclick="GlpiModule.showMore()">Mostrar ${Math.min(PAGE_SIZE, tickets.length - visible.length)} más</button>
           <span>Mostrando ${visible.length.toLocaleString('es-CO')} de ${tickets.length.toLocaleString('es-CO')}</span>`
        : `<span>Mostrando ${visible.length.toLocaleString('es-CO')} tickets</span>`;
    }
  }

  function parseDescriptionComponents(description){
    let source = cleanPlainText(description)
      .replace(/Datos del formulario/gi, '')
      .replace(/Section\s*/gi, '')
      .trim();
    if(!source) return [];
    const starts = [];
    let cursor = 0;
    for(let number = 1; number <= 50; number += 1) {
      const marker = `${number})`;
      const index = source.indexOf(marker, cursor);
      if(index < 0) {
        if(number === 1) break;
        continue;
      }
      starts.push({ number, index, markerLength:marker.length });
      cursor = index + marker.length;
    }
    if(!starts.length) return [];
    return starts.map((start, index) => {
      const next = starts[index + 1];
      const segment = source.slice(start.index + start.markerLength, next ? next.index : source.length).trim();
      const separator = segment.indexOf(':');
      return {
        label: separator >= 0
          ? segment.slice(0, separator).replace(/^[\s.:\-]+|[\s.:\-]+$/g, '').trim()
          : `Campo ${start.number}`,
        value: separator >= 0
          ? segment.slice(separator + 1).replace(/^[\s.:\-]+/, '').trim()
          : segment
      };
    }).filter(item => item.label || item.value);
  }

  function renderDescriptionContent(description, components){
    if(!components.length) {
      return `<p class="glpi-description-plain glpi-description-full">${escapeHtml(description || 'Sin descripción registrada')}</p>`;
    }
    return `<ol class="glpi-description-list">${components.map((component, index) => `<li>
      <span class="glpi-description-number">${index + 1}</span>
      <div>
        <strong>${escapeHtml(component.label || `Campo ${index + 1}`)}</strong>
        <p>${escapeHtml(component.value || 'Sin información')}</p>
      </div>
    </li>`).join('')}</ol>`;
  }

  function renderDetail(ticket){
    const host = document.getElementById('glpi-detail-content');
    if(!host) return;
    const components = parseDescriptionComponents(ticket.description);
    const time = getTicketTimeInfo(ticket, getContextMetrics());
    host.innerHTML = `
      <div class="glpi-detail-eyebrow">Ticket #${escapeHtml(ticket.id)}</div>
      <h2 id="glpi-detail-title">${escapeHtml(ticket.title)}</h2>
      <div class="glpi-detail-status-line">
        <span class="glpi-status glpi-status-${getStatusClass(ticket.statusKey)}">${escapeHtml(ticket.status)}</span>
        <span>${escapeHtml(ticket.directorGroup)}</span>
      </div>
      <div class="glpi-detail-scroll-cue"><span>Detalle completo del ticket</span><span>Desplázate para ver todos los campos ↓</span></div>
      <div class="glpi-detail-metrics">
        <div><span>Apertura</span><strong>${escapeHtml(formatDate(ticket.openDate))}</strong></div>
        <div><span>Solución</span><strong>${escapeHtml(formatDate(ticket.solutionDate))}</strong></div>
        <div><span>Cierre</span><strong>${escapeHtml(formatDate(ticket.closeDate))}</strong></div>
        <div><span>${escapeHtml(time.label)}</span><strong>${escapeHtml(formatDaysWithUnit(time.days))}</strong></div>
        <div><span>Vs. promedio del contexto</span><strong class="glpi-time-${time.comparisonClass}">${escapeHtml(time.comparison)}</strong></div>
      </div>
      <div class="glpi-detail-info">
        <div><span>Solicitante</span><strong>${escapeHtml(ticket.requester)}</strong></div>
        <div><span>Grupo del director</span><strong>${escapeHtml(ticket.directorGroup)}</strong></div>
        <div><span>Técnico asignado</span><strong>${escapeHtml(ticket.technicians.join(', ') || 'Sin asignar')}</strong></div>
        <div><span>Categoría completa</span><strong>${escapeHtml(ticket.category)}</strong></div>
      </div>
      <section class="glpi-description-section">
        <h3>Descripción completa del caso</h3>
        ${renderDescriptionContent(ticket.description, components)}
      </section>
      <div class="glpi-detail-end"><span>✓</span> Fin del detalle del ticket</div>`;
  }

  function renderSourceNote(){
    const host = document.getElementById('glpi-source-note');
    if(!host) return;
    host.textContent = `Fuente: ${GLPI_FILE_NAME} · hoja ${GLPI_SHEET_NAME} · actualizado ${formatSourceDate(state.sourceModifiedAt)} · ${state.tickets.length.toLocaleString('es-CO')} tickets únicos · ${state.duplicateRows.toLocaleString('es-CO')} filas duplicadas consolidadas.`;
  }

  function renderAll(){
    ensureDefaultPeriod();
    renderPeriodOptions();
    if(state.loading) {
      renderLoading();
      return;
    }
    if(state.error) {
      renderError();
      return;
    }
    if(!state.loaded) return;
    renderFilterOptions();
    renderSourceNote();
    renderContextSummary();
    renderKpis();
    renderStatusDonutChart(getContextMetrics());
    renderGroups();
    renderTable();
  }

  async function resolveDownloadUrl(item, token){
    if(item && item['@microsoft.graph.downloadUrl']) return item['@microsoft.graph.downloadUrl'];
    const driveId = item && item.parentReference && item.parentReference.driveId;
    if(!driveId || !item.id) return '';
    const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}?$select=id,name,lastModifiedDateTime,@microsoft.graph.downloadUrl`, {
      headers: { Authorization:`Bearer ${token}` }
    });
    const payload = await response.json().catch(() => ({}));
    if(!response.ok) throw new Error((payload.error && payload.error.message) || 'No se pudo descargar el archivo GLPI.');
    return payload['@microsoft.graph.downloadUrl'] || '';
  }

  async function loadFromSharePoint(){
    if(state.loading) return;
    state.loading = true;
    state.error = '';
    renderAll();
    try {
      if(typeof XLSX === 'undefined') throw new Error('El lector de Excel no está disponible.');
      const siteId = await getSiteId();
      const token = await getToken(['Files.Read.All']);
      const basePath = await getForecastBasePath(siteId, token);
      const folderPath = joinGraphPath(basePath, GLPI_FOLDER_NAME);
      const { r, d } = await fetchGraphPathJson(siteId, folderPath, 'children?$top=50', token);
      if(!r.ok) throw new Error((d.error && d.error.message) || 'No se pudo abrir la carpeta GLPI.');
      const item = (d.value || []).find(candidate =>
        candidate && normalizeText(candidate.name) === normalizeText(GLPI_FILE_NAME)
      );
      if(!item) throw new Error(`No se encontró ${GLPI_FILE_NAME} en la carpeta GLPI.`);
      const downloadUrl = await resolveDownloadUrl(item, token);
      if(!downloadUrl) throw new Error(`SharePoint no entregó una URL de descarga para ${GLPI_FILE_NAME}.`);
      const fileResponse = await fetch(downloadUrl, { cache:'no-store' });
      if(!fileResponse.ok) throw new Error(`No se pudo descargar ${GLPI_FILE_NAME} (HTTP ${fileResponse.status}).`);
      const buffer = await fileResponse.arrayBuffer();
      const workbook = XLSX.read(buffer, { type:'array', cellDates:true });
      const sheetName = (workbook.SheetNames || []).find(name => normalizeText(name) === normalizeText(GLPI_SHEET_NAME));
      if(!sheetName) throw new Error(`El Excel ${GLPI_FILE_NAME} no contiene la hoja ${GLPI_SHEET_NAME}.`);
      const sheet = workbook.Sheets[sheetName];
      if(!sheet) throw new Error(`La hoja ${GLPI_SHEET_NAME} de ${GLPI_FILE_NAME} no se puede leer.`);
      const rows = XLSX.utils.sheet_to_json(sheet, { defval:null, raw:true });
      const tickets = deduplicateTickets(rows);
      if(!tickets.length) throw new Error('El Excel TICKETS-GLPI no contiene tickets válidos.');
      state.tickets = tickets;
      state.sourceRows = rows.length;
      state.duplicateRows = Math.max(0, rows.length - tickets.length);
      state.sourceModifiedAt = item.lastModifiedDateTime || '';
      state.loaded = true;
      state.limit = PAGE_SIZE;
      ensureDefaultPeriod();
    } catch(error) {
      console.error('[GLPI]', error);
      state.error = error instanceof Error ? error.message : 'No fue posible cargar TICKETS-GLPI.';
    } finally {
      state.loading = false;
      renderAll();
    }
  }

  function render(){
    if(!state.loaded && !state.loading && !state.error) {
      loadFromSharePoint();
      return;
    }
    renderAll();
  }

  function resetPaging(){
    state.limit = PAGE_SIZE;
  }

  function setPeriod(value){
    state.period = value || '';
    state.group = '';
    state.requester = '';
    resetPaging();
    renderAll();
  }

  function setStatus(value){
    state.status = ['all','new','in_progress','waiting','resolved','closed'].includes(value) ? value : 'all';
    resetPaging();
    renderAll();
  }

  function toggleStatus(value){
    const selected = ['all','new','in_progress','waiting','resolved','closed'].includes(value) ? value : 'all';
    state.status = selected !== 'all' && state.status === selected ? 'all' : selected;
    resetPaging();
    renderAll();
  }

  function setGroup(value){
    state.group = value || '';
    resetPaging();
    renderAll();
  }

  function toggleGroup(value){
    state.group = state.group === value ? '' : (value || '');
    resetPaging();
    renderAll();
  }

  function setRequester(value){
    state.requester = value || '';
    resetPaging();
    renderAll();
  }

  function toggleRequester(value){
    state.requester = state.requester === value ? '' : (value || '');
    resetPaging();
    renderAll();
  }

  function setSearch(value){
    state.search = value || '';
    resetPaging();
    renderAll();
  }

  function clearFilters(){
    state.status = 'all';
    state.group = '';
    state.requester = '';
    state.search = '';
    resetPaging();
    renderAll();
  }

  function showMore(){
    state.limit += PAGE_SIZE;
    renderTable();
  }

  function openTicket(id){
    const ticket = state.tickets.find(item => item.id === String(id));
    if(!ticket) return;
    state.selectedId = ticket.id;
    renderDetail(ticket);
    const backdrop = document.getElementById('glpi-detail-backdrop');
    const panel = document.getElementById('glpi-detail-panel');
    if(panel) panel.scrollTop = 0;
    if(backdrop) {
      backdrop.classList.add('open');
      backdrop.setAttribute('aria-hidden', 'false');
    }
    document.body.classList.add('glpi-detail-open');
    const closeButton = panel && panel.querySelector ? panel.querySelector('.glpi-detail-close') : null;
    if(closeButton && typeof closeButton.focus === 'function') closeButton.focus({ preventScroll:true });
  }

  function closeDetail(event){
    if(event && event.target && event.currentTarget && event.target !== event.currentTarget) return;
    state.selectedId = '';
    const backdrop = document.getElementById('glpi-detail-backdrop');
    if(backdrop) {
      backdrop.classList.remove('open');
      backdrop.setAttribute('aria-hidden', 'true');
    }
    document.body.classList.remove('glpi-detail-open');
  }

  async function reload(){
    state.loaded = false;
    state.error = '';
    state.tickets = [];
    await loadFromSharePoint();
  }

  document.addEventListener('keydown', event => {
    if(event.key === 'Escape' && state.selectedId) closeDetail();
  });

  window.GlpiModule = {
    render,
    reload,
    loadFromSharePoint,
    setPeriod,
    setStatus,
    toggleStatus,
    setGroup,
    toggleGroup,
    setRequester,
    toggleRequester,
    setSearch,
    clearFilters,
    showMore,
    openTicket,
    closeDetail
  };
  ensureDefaultPeriod();
  renderPeriodOptions();
})();
