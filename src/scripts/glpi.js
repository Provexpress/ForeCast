(function(){
  'use strict';

  const GLPI_FOLDER_NAME = 'GLPI';
  const GLPI_FILE_NAME = 'TICKETS-GLPI.xlsx';
  const GLPI_SHEET_NAME = '2026';
  const PAGE_SIZE = 60;
  let commercialPeopleCache = null;

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

  function diffDays(startIso, endIso){
    if(!startIso) return null;
    const start = new Date(`${startIso}T00:00:00Z`);
    const end = new Date(`${endIso || new Date().toISOString().slice(0, 10)}T00:00:00Z`);
    if(Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return null;
    return Math.max(0, Math.round((end - start) / 86400000));
  }

  function getStatusKey(value){
    const normalized = normalizeText(value);
    if(normalized.includes('nuevo')) return 'new';
    if(normalized.includes('cerrad') || normalized.includes('resuelt') || normalized.includes('solucion')) return 'closed';
    return 'open';
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
    const status = cleanPlainText(getField(lookup, ['Estado'])) || 'Sin estado';
    const statusKey = getStatusKey(status);
    const requester = cleanPersonName(getField(lookup, ['Solicitante - Solicitante', 'Solicitante'])) || 'Sin solicitante';
    const category = cleanPlainText(getField(lookup, ['Categoría', 'Categoria'])) || 'Sin categoría';
    const directorGroup = resolveDirectorGroup(requester);
    const technicians = splitPeople(getField(lookup, ['Asignado a: - Técnico', 'Asignado a - Técnico', 'Técnico']));
    const sourceDaysSolution = parseNumber(getField(lookup, ['Días en Solución', 'Dias en Solucion']));
    const sourceDaysOpen = parseNumber(getField(lookup, ['Días Abierto', 'Dias Abierto']));
    const daysSolution = sourceDaysSolution === null ? diffDays(openDate, solutionDate) : sourceDaysSolution;
    const daysOpen = statusKey === 'closed'
      ? null
      : (sourceDaysOpen === null ? diffDays(openDate) : sourceDaysOpen);

    return {
      id: id || `fila-${index + 1}`,
      title: cleanPlainText(getField(lookup, ['Título', 'Titulo'])) || 'Ticket sin título',
      status,
      statusKey,
      openDate,
      solutionDate,
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
      candidate.category,
      candidate.description,
      candidate.requester
    ].filter(Boolean).length > [
      current.title,
      current.status,
      current.openDate,
      current.solutionDate,
      current.category,
      current.description,
      current.requester
    ].filter(Boolean).length;
    const preferred = preferCandidate ? candidate : current;
    Object.assign(merged, preferred);
    merged.technicians = [...new Set([...(current.technicians || []), ...(candidate.technicians || [])])];
    if((candidate.description || '').length > (current.description || '').length) merged.description = candidate.description;
    if(current.daysSolution !== null && current.daysSolution !== undefined) merged.daysSolution = current.daysSolution;
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
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone:'America/Bogota',
      year:'numeric',
      month:'2-digit'
    }).formatToParts(new Date()).reduce((acc, part) => {
      if(part.type !== 'literal') acc[part.type] = part.value;
      return acc;
    }, {});
    return `${parts.year}-${parts.month}`;
  }

  function getPeriods(){
    return [...new Set([
      getCurrentPeriod(),
      ...state.tickets.map(ticket => ticket.period)
    ].filter(period => /^\d{4}-\d{2}$/.test(String(period || ''))))].sort().reverse();
  }

  function ensureDefaultPeriod(){
    const periods = getPeriods();
    if(state.period && periods.includes(state.period)) return;
    state.period = getCurrentPeriod();
  }

  function formatMonth(period){
    if(!period) return 'Sin mes';
    const [year, month] = period.split('-').map(Number);
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

  function average(values){
    const valid = values.map(Number).filter(Number.isFinite);
    return valid.length ? valid.reduce((sum, value) => sum + value, 0) / valid.length : null;
  }

  function getMonthTickets(){
    return state.tickets.filter(ticket => !state.period || ticket.period === state.period);
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

  function getFilteredTickets(){
    return getContextTickets().filter(ticket => state.status === 'all' || ticket.statusKey === state.status);
  }

  function getStatusLabel(key){
    return ({ new:'Nuevo', open:'Abierto', closed:'Cerrado' })[key] || 'Sin estado';
  }

  function getStatusClass(key){
    return ({ new:'new', open:'open', closed:'closed' })[key] || 'open';
  }

  function renderLoading(){
    const kpis = document.getElementById('glpi-kpis');
    const table = document.getElementById('glpi-ticket-table');
    const source = document.getElementById('glpi-source-note');
    if(source) source.textContent = 'Conectando con TICKETS-GLPI.xlsx en SharePoint…';
    if(kpis) kpis.innerHTML = '<div class="glpi-loading-card">Leyendo y organizando tickets…</div>';
    if(table) table.innerHTML = '<div class="glpi-empty-state">Preparando datos del Excel.</div>';
  }

  function renderError(){
    const message = state.error || 'No fue posible cargar los tickets.';
    const source = document.getElementById('glpi-source-note');
    const kpis = document.getElementById('glpi-kpis');
    const table = document.getElementById('glpi-ticket-table');
    if(source) source.textContent = message;
    if(kpis) kpis.innerHTML = `<div class="glpi-error-card">${escapeHtml(message)}</div>`;
    if(table) table.innerHTML = '<div class="glpi-empty-state">Usa “Actualizar Excel” para intentar de nuevo.</div>';
  }

  function renderPeriodOptions(){
    const select = document.getElementById('glpi-period');
    if(!select) return;
    select.innerHTML = getPeriods()
      .map(period => `<option value="${escapeAttr(period)}"${period === state.period ? ' selected' : ''}>${escapeHtml(formatMonth(period))}</option>`)
      .join('');
  }

  function renderFilterOptions(){
    const monthTickets = getMonthTickets();
    const groups = [...new Set(monthTickets.map(ticket => ticket.directorGroup).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    const requesters = [...new Set(monthTickets.map(ticket => ticket.requester).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    const groupSelect = document.getElementById('glpi-group-filter');
    const requesterSelect = document.getElementById('glpi-requester-filter');
    const statusSelect = document.getElementById('glpi-status-filter');
    const searchInput = document.getElementById('glpi-search');

    if(state.group && !groups.includes(state.group)) state.group = '';
    if(state.requester && !requesters.includes(state.requester)) state.requester = '';
    if(groupSelect) {
      groupSelect.innerHTML = '<option value="">Todos los grupos de directores</option>' + groups
        .map(group => `<option value="${escapeAttr(group)}"${group === state.group ? ' selected' : ''}>${escapeHtml(group)}</option>`)
        .join('');
    }
    if(requesterSelect) {
      requesterSelect.innerHTML = '<option value="">Todos los solicitantes</option>' + requesters
        .map(requester => `<option value="${escapeAttr(requester)}"${requester === state.requester ? ' selected' : ''}>${escapeHtml(requester)}</option>`)
        .join('');
    }
    if(statusSelect) statusSelect.value = state.status;
    if(searchInput && searchInput.value !== state.search) searchInput.value = state.search;
  }

  function kpiCard(kind, label, value, detail){
    const active = state.status === kind;
    const clickable = ['new','open','closed'].includes(kind);
    const tag = clickable ? 'button' : 'article';
    const action = clickable ? ` type="button" onclick="GlpiModule.setStatus('${kind}')"` : '';
    return `<${tag} class="glpi-kpi-card glpi-kpi-${escapeAttr(kind)}${active ? ' active' : ''}"${action}>
      <span class="glpi-kpi-label">${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <small>${escapeHtml(detail)}</small>
    </${tag}>`;
  }

  function renderKpis(){
    const host = document.getElementById('glpi-kpis');
    if(!host) return;
    const context = getContextTickets();
    const newTickets = context.filter(ticket => ticket.statusKey === 'new');
    const openTickets = context.filter(ticket => ticket.statusKey === 'open');
    const closedTickets = context.filter(ticket => ticket.statusKey === 'closed');
    const avgSolution = average(closedTickets.map(ticket => ticket.daysSolution));
    const avgOpen = average([...newTickets, ...openTickets].map(ticket => ticket.daysOpen));
    host.innerHTML = [
      kpiCard('total', 'Tickets del mes', context.length.toLocaleString('es-CO'), formatMonth(state.period)),
      kpiCard('new', 'Nuevos', newTickets.length.toLocaleString('es-CO'), 'Sin iniciar'),
      kpiCard('open', 'Abiertos', openTickets.length.toLocaleString('es-CO'), 'En gestión'),
      kpiCard('closed', 'Cerrados', closedTickets.length.toLocaleString('es-CO'), 'Resueltos o cerrados'),
      kpiCard('solution', 'Días de solución', formatDays(avgSolution), 'Promedio de cerrados'),
      kpiCard('age', 'Días abiertos', formatDays(avgOpen), 'Promedio pendiente')
    ].join('');
  }

  function summarizeBy(tickets, getter){
    const groups = new Map();
    tickets.forEach(ticket => {
      const name = getter(ticket) || 'Sin grupo';
      if(!groups.has(name)) groups.set(name, { name, total:0, new:0, open:0, closed:0, days:[] });
      const group = groups.get(name);
      group.total += 1;
      group[ticket.statusKey] += 1;
      const days = ticket.statusKey === 'closed' ? ticket.daysSolution : ticket.daysOpen;
      if(Number.isFinite(Number(days))) group.days.push(Number(days));
    });
    return [...groups.values()]
      .map(group => ({ ...group, averageDays: average(group.days) }))
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
      const action = filterType === 'group'
        ? `GlpiModule.setGroup(${jsString(group.name)})`
        : `GlpiModule.setRequester(${jsString(group.name)})`;
      return `<button class="glpi-rank-row" type="button" onclick="${escapeAttr(action)}">
        <span class="glpi-rank-main">
          <span class="glpi-rank-name" title="${escapeAttr(group.name)}">${escapeHtml(group.name)}</span>
          <span class="glpi-rank-value">${group.total.toLocaleString('es-CO')}</span>
        </span>
        <span class="glpi-rank-track"><span style="width:${Math.max(4, group.total / max * 100).toFixed(1)}%"></span></span>
        <span class="glpi-rank-meta">
          <span>${group.open + group.new} pendientes</span>
          <span>${group.closed} cerrados</span>
          <span>${formatDays(group.averageDays)} días prom.</span>
        </span>
      </button>`;
    }).join('');
  }

  function renderGroups(){
    const tickets = getFilteredTickets();
    renderRankedList('glpi-category-groups', summarizeBy(tickets, ticket => ticket.directorGroup), 'group');
    renderRankedList('glpi-requester-groups', summarizeBy(tickets, ticket => ticket.requester), 'requester');
  }

  function renderTable(){
    const host = document.getElementById('glpi-ticket-table');
    const meta = document.getElementById('glpi-table-meta');
    const more = document.getElementById('glpi-table-more');
    if(!host) return;
    const tickets = getFilteredTickets();
    const visible = tickets.slice(0, state.limit);
    if(meta) meta.textContent = `${tickets.length.toLocaleString('es-CO')} tickets después de filtros · clic en una fila para ver el detalle`;
    if(!visible.length) {
      host.innerHTML = '<div class="glpi-empty-state">No hay tickets para los filtros seleccionados.</div>';
      if(more) more.innerHTML = '';
      return;
    }

    host.innerHTML = `<table class="responsive-table glpi-table">
      <thead><tr>
        <th>ID</th><th>Título</th><th>Estado</th><th>Apertura</th><th>Grupo director</th>
        <th>Categoría</th><th>Solicitante</th><th>Técnico</th><th>Días solución</th><th>Días abierto</th>
      </tr></thead>
      <tbody>${visible.map(ticket => `<tr class="glpi-ticket-row" onclick="GlpiModule.openTicket('${escapeAttr(ticket.id)}')" tabindex="0" onkeydown="if(event.key==='Enter')GlpiModule.openTicket('${escapeAttr(ticket.id)}')">
        <td class="td-mono" data-label="ID">#${escapeHtml(ticket.id)}</td>
        <td class="glpi-title-cell" data-label="Título" title="${escapeAttr(ticket.title)}">${escapeHtml(ticket.title)}</td>
        <td data-label="Estado"><span class="glpi-status glpi-status-${getStatusClass(ticket.statusKey)}">${escapeHtml(ticket.status)}</span></td>
        <td class="td-mono" data-label="Apertura">${escapeHtml(formatDate(ticket.openDate))}</td>
        <td data-label="Grupo director">${escapeHtml(ticket.directorGroup)}</td>
        <td data-label="Categoría">${escapeHtml(ticket.categoryGroup)}</td>
        <td data-label="Solicitante">${escapeHtml(ticket.requester)}</td>
        <td data-label="Técnico">${escapeHtml(ticket.technicians.join(', ') || 'Sin asignar')}</td>
        <td class="td-mono" data-label="Días solución">${escapeHtml(formatDays(ticket.daysSolution))}</td>
        <td class="td-mono" data-label="Días abierto">${escapeHtml(formatDays(ticket.daysOpen))}</td>
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
        label: separator >= 0 ? segment.slice(0, separator).trim() : `Campo ${start.number}`,
        value: separator >= 0 ? segment.slice(separator + 1).trim() : segment
      };
    }).filter(item => item.label || item.value);
  }

  function renderDetail(ticket){
    const host = document.getElementById('glpi-detail-content');
    if(!host) return;
    const components = parseDescriptionComponents(ticket.description);
    host.innerHTML = `
      <div class="glpi-detail-eyebrow">Ticket #${escapeHtml(ticket.id)}</div>
      <h2 id="glpi-detail-title">${escapeHtml(ticket.title)}</h2>
      <div class="glpi-detail-status-line">
        <span class="glpi-status glpi-status-${getStatusClass(ticket.statusKey)}">${escapeHtml(ticket.status)}</span>
        <span>${escapeHtml(ticket.directorGroup)}</span>
      </div>
      <div class="glpi-detail-metrics">
        <div><span>Apertura</span><strong>${escapeHtml(formatDate(ticket.openDate))}</strong></div>
        <div><span>Solución</span><strong>${escapeHtml(formatDate(ticket.solutionDate))}</strong></div>
        <div><span>Días solución</span><strong>${escapeHtml(formatDays(ticket.daysSolution))}</strong></div>
        <div><span>Días abierto</span><strong>${escapeHtml(formatDays(ticket.daysOpen))}</strong></div>
      </div>
      <div class="glpi-detail-info">
        <div><span>Solicitante</span><strong>${escapeHtml(ticket.requester)}</strong></div>
        <div><span>Grupo del director</span><strong>${escapeHtml(ticket.directorGroup)}</strong></div>
        <div><span>Técnico asignado</span><strong>${escapeHtml(ticket.technicians.join(', ') || 'Sin asignar')}</strong></div>
        <div><span>Categoría completa</span><strong>${escapeHtml(ticket.category)}</strong></div>
      </div>
      <section class="glpi-description-section">
        <h3>Descripción y componentes</h3>
        ${components.length ? `<div class="glpi-component-list">${components.map(component => `<article>
          <span>${escapeHtml(component.label || 'Campo')}</span>
          <p>${escapeHtml(component.value || 'Sin información')}</p>
        </article>`).join('')}</div>` : `<p class="glpi-description-plain">${escapeHtml(ticket.description || 'Sin descripción')}</p>`}
      </section>`;
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
    renderKpis();
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
    state.status = ['all','new','open','closed'].includes(value) ? value : 'all';
    resetPaging();
    renderAll();
  }

  function setGroup(value){
    state.group = value || '';
    resetPaging();
    renderAll();
  }

  function setRequester(value){
    state.requester = value || '';
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
    if(backdrop) {
      backdrop.classList.add('open');
      backdrop.setAttribute('aria-hidden', 'false');
    }
    document.body.classList.add('glpi-detail-open');
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
    setGroup,
    setRequester,
    setSearch,
    clearFilters,
    showMore,
    openTicket,
    closeDetail
  };
  ensureDefaultPeriod();
  renderPeriodOptions();
})();
