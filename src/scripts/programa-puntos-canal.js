/* ══════════════════════════════════════
   REBATES, PUNTOS E INCENTIVOS POR CANAL
   Normalizacion de reportes de fabricantes
══════════════════════════════════════ */
(function(){
  'use strict';

  const SHAREPOINT_FOLDER_NAME = 'Fabricantes';

  const SOURCE_DEFINITIONS = [
    {
      key: 'platforms',
      label: 'Reporte plataformas',
      queries: ['Reporte palataformas Juan 2026', 'Reporte plataformas Juan 2026'],
      matches: name => /reporte (palataformas|plataformas) juan 2026/.test(name)
    },
    {
      key: 'myrewards',
      label: 'MyRewards Dell',
      queries: ['MyRewards Partner Detail Report'],
      matches: name => name.includes('myrewards partner detail report')
    },
    {
      key: 'claims',
      label: 'Claim Summary Dell',
      queries: ['Claim Summary Report'],
      matches: name => name.includes('claim summary report')
    },
    {
      key: 'lenovo',
      label: 'Program List Lenovo',
      queries: ['Program_List_LAS_PA000006347092', 'Program List LAS PA000006347092'],
      matches: name => name.includes('program list las pa000006347092')
    }
  ];

  const CHANNELS = ['Dell','Lenovo','HPE','ASUS','Epson','Intel'];
  const VALIDATED_AT = '14 jul 2026';
  const BUSINESS_METRICS = Object.freeze({
    lenovo: {
      actual2026: 20633.90,
      projection2026: 82535.60,
      yoy2026: 21,
      history: { FY23:19962.95, FY24:58529.23, FY25:68165.62 }
    },
    dell: { approved:49053.25, unredeemedPoints:265 },
    asus: { quota:200000, sales:38472, units:60, compliance:19 },
    hpe: { availablePoints:390, visaUsd:300 },
    sed: { travelerBonusCop:5198820, validityMonths:5 }
  });

  const VALIDATED_RECORD_DEFINITIONS = [
    // Lenovo 360 Engage — histórico anual y programas FY-Q sin inventar distribuciones.
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'Lenovo 360 Engage · CASH', periodo:'FY23', valor:19962.95, unidad:'USD', estado:'Pagada', clienteRef:'Crecimiento base 2023' },
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'Lenovo 360 Engage · CASH', periodo:'FY24', valor:58529.23, unidad:'USD', estado:'Pagada', clienteRef:'+193% YoY' },
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'Lenovo 360 Engage · CASH', periodo:'FY25', valor:68165.62, unidad:'USD', estado:'Pagada', clienteRef:'+16% YoY · estabilización' },
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'Lenovo 360 Engage · CASH', periodo:'FY26', valor:20633.90, unidad:'USD', estado:'Pendiente', clienteRef:'Actual 2026 · proyección USD 82.535,60' },
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'Engage Platinum', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'Pendiente', clienteRef:'Modalidad CASH · T2' },
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'Advocate', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'Pendiente', clienteRef:'Modalidad CASH · T2' },
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'WKS Expert', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'Oportunidad', clienteRef:'Workstation subexplotado · T2' },
    { grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', programa:'BDF Platinum', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'Pendiente', clienteRef:'Modalidad CASH · T2' },

    // Dell MDF / Rebates — Approved to Pay validado.
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Base', periodo:'FY26-Q1', valor:5532.54, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'eMDF', periodo:'FY26-Q1', valor:1871.82, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Services', periodo:'FY26-Q1', valor:420.29, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'eMDF (FY25)', periodo:'FY26-Q2', valor:1930.78, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q2' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Base', periodo:'FY26-Q2', valor:5480.99, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q2' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'eMDF', periodo:'FY26-Q2', valor:1841.93, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q2' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'NBI', periodo:'FY26-Q2', valor:1405.08, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q2' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Services', periodo:'FY26-Q2', valor:807.68, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q2' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'eMDF (FY25)', periodo:'FY26-Q3', valor:1338.84, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q3' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Base', periodo:'FY26-Q3', valor:4587.57, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q3' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'NBI', periodo:'FY26-Q3', valor:2322, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q3' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Services', periodo:'FY26-Q3', valor:1580.75, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q3' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'eMDF (FY25)', periodo:'FY26-Q4', valor:2471.74, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q4' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Base', periodo:'FY26-Q4', valor:1917.30, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q4' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'pbMDF ISG', periodo:'FY26-Q4', valor:2000, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q4' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Services', periodo:'FY26-Q4', valor:483.03, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY26 Q4' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Services Cloud Service Provider', periodo:'FY26-Q4', valor:0, unidad:'USD', estado:'Declined', clienteRef:'FY26 Q4' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Base Cloud', periodo:'FY27-Q1', valor:0, unidad:'USD', estado:'Pending Final Review', clienteRef:'FY27 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Base SP', periodo:'FY27-Q1', valor:8364.83, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY27 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'pbMDF Client+', periodo:'FY27-Q1', valor:0, unidad:'USD', estado:'Released', clienteRef:'FY27 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'pbMDF ISG', periodo:'FY27-Q1', valor:0, unidad:'USD', estado:'Released', clienteRef:'FY27 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Services Cloud', periodo:'FY27-Q1', valor:0, unidad:'USD', estado:'Pending Final Review', clienteRef:'FY27 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Services SP', periodo:'FY27-Q1', valor:1289.08, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY27 Q1' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'eMDF', periodo:'FY27-Q3', valor:0, unidad:'USD', estado:'Released', clienteRef:'FY27 Q3' },
    { grupo:'claims', canal:'Dell', tipo:'Rebate', programa:'Base SP', periodo:'FY27-Q4', valor:3407, unidad:'USD', estado:'Paid/Complete', clienteRef:'FY27 Q4' },

    // Canales sin rebate liquidado informado: se registra el programa, no se inventa ingreso.
    { grupo:'asus', canal:'ASUS', tipo:'Rebate', programa:'AGP Program 2026 · AGP Silver', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'Alerta · 19%', clienteRef:'KAM Oscar Bolaños · 60 unidades · NEXSYS, INGRAM, MPS, IMPRESISTEM' },
    { grupo:'asus', canal:'ASUS', tipo:'Rebate', programa:'AGP Silver 2025 · Rebate 0–0.5', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'MDF Elegible', clienteRef:'Fuerza de ventas: 1 cupo · Excluye Vivobook, Zenbook, ROG, TUF, monitores y Mini PC/NUC' },
    { grupo:'platforms', canal:'Epson', tipo:'Rebate', programa:'Consumo (Colombia)', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'Pendiente', clienteRef:'Mayoristas: SED, Ingram, Impresistem, Nexsys' },
    { grupo:'platforms', canal:'Epson', tipo:'Rebate', programa:'Comercial (Colombia)', periodo:'FY26-Q2', valor:0, unidad:'USD', estado:'Procesado', clienteRef:'Mayoristas: SED, Ingram, Impresistem, Nexsys' },

    // Dell MyRewards — histórico y saldo pendiente de redención.
    { grupo:'myrewards', metrica:'movement', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards', periodo:'FY22-Q1', valor:50, unidad:'Puntos', estado:'Allocated', clienteRef:'Histórico trimestral' },
    { grupo:'myrewards', metrica:'movement', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards', periodo:'FY22-Q2', valor:415, unidad:'Puntos', estado:'Allocated', clienteRef:'Histórico trimestral' },
    { grupo:'myrewards', metrica:'movement', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards', periodo:'FY22-Q3', valor:540, unidad:'Puntos', estado:'Allocated', clienteRef:'Histórico trimestral' },
    { grupo:'myrewards', metrica:'movement', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards', periodo:'FY22-Q4', valor:375, unidad:'Puntos', estado:'Allocated', clienteRef:'Histórico trimestral' },
    { grupo:'myrewards', metrica:'movement', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards', periodo:'FY23-Q1', valor:135, unidad:'Puntos', estado:'Allocated', clienteRef:'Histórico trimestral' },
    { grupo:'myrewards', metrica:'movement', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards', periodo:'FY23-Q2', valor:120, unidad:'Puntos', estado:'Allocated', clienteRef:'Histórico trimestral' },
    { grupo:'myrewards', metrica:'balance', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards · Pendiente de redención', periodo:'Sin periodo', valor:265, unidad:'Puntos', estado:'Unclaimed', clienteRef:'Acción requerida por gerencia' },
    { grupo:'myrewards', metrica:'balance', canal:'Dell', tipo:'Punto', programa:'Dell MyRewards · Expired', periodo:'Sin periodo', valor:0, unidad:'Puntos', estado:'Expired', clienteRef:'Estado disponible para seguimiento' },

    // HPE Instant On e incentivos SED.
    { grupo:'platforms', metrica:'balance', canal:'HPE', tipo:'Punto', programa:'HPE Instant On · Saldo disponible', periodo:'FY26-Q2', valor:390, unidad:'Puntos', estado:'Disponible', clienteRef:'Visa Prepaid Card USD 300' },
    { grupo:'platforms', metrica:'movement', canal:'HPE', tipo:'Punto', programa:'HPE Instant On · Cash bonus', periodo:'FY25-Q1', valor:215, unidad:'Puntos', estado:'Allocated', clienteRef:'Cash bonus trimestral' },
    { grupo:'platforms', metrica:'movement', canal:'HPE', tipo:'Punto', programa:'HPE Instant On · Cash bonus', periodo:'FY25-Q3', valor:403, unidad:'Puntos', estado:'Allocated', clienteRef:'Cash bonus trimestral' },
    { grupo:'platforms', metrica:'movement', canal:'HPE', tipo:'Punto', programa:'HPE Instant On · Cash bonus', periodo:'FY25-Q4', valor:88, unidad:'Puntos', estado:'Allocated', clienteRef:'Cash bonus trimestral' },
    { grupo:'platforms', metrica:'movement', canal:'HPE', tipo:'Punto', programa:'HPE Instant On · Cash bonus', periodo:'FY26-Q1', valor:95, unidad:'Puntos', estado:'Allocated', clienteRef:'Cash bonus trimestral' },
    { grupo:'platforms', metrica:'movement', canal:'HPE', tipo:'Punto', programa:'HPE Instant On · Cash bonus', periodo:'FY26-Q2', valor:242, unidad:'Puntos', estado:'Allocated', clienteRef:'Cash bonus trimestral' },
    { grupo:'sed', metrica:'movement', canal:'Lenovo', tipo:'Punto', programa:'Plan Ultra Aguinaldo Lenovo WS', periodo:'Sin periodo', valor:37, unidad:'Puntos', estado:'Disponible', clienteRef:'Incentivo SED · vigencia 5 meses' },
    { grupo:'sed', metrica:'movement', canal:'HPE', tipo:'Punto', programa:'HPE Speed Month', periodo:'Sin periodo', valor:259, unidad:'Puntos', estado:'Disponible', clienteRef:'Incentivo SED · vigencia 5 meses' },
    { grupo:'sed', metrica:'movement', canal:'HPE', tipo:'Punto', programa:'Multiplica tus Ingresos HPE Networking', periodo:'Sin periodo', valor:210, unidad:'Puntos', estado:'Disponible', clienteRef:'Incentivo SED · vigencia 5 meses' },
    { grupo:'sed', metrica:'balance', canal:'HPE', tipo:'Punto', programa:'Escalera de Premios HPE Networking', periodo:'Sin periodo', valor:0, unidad:'Puntos', estado:'Cargado a tarjeta Big Pass', clienteRef:'Incentivo SED · vigencia 5 meses' },
    { grupo:'sed', metrica:'balance', canal:'Lenovo', tipo:'Punto', programa:'Bono Viajero Lenovo Adventure', periodo:'Sin periodo', valor:5198820, unidad:'COP', estado:'Vigente', clienteRef:'Incentivo SED · vigencia 5 meses' },
    { grupo:'platforms', metrica:'balance', canal:'Intel', tipo:'Punto', programa:'Programa de incentivos Intel', periodo:'Sin periodo', valor:0, unidad:'Puntos', estado:'Pendiente', clienteRef:'Sin valor confirmado' }
  ];

  const state = {
    mode: 'Rebate',
    filters: { canal:'', periodo:'', estado:'', unidad:'USD' },
    page: 1,
    pageSize: 20,
    loading: false,
    error: '',
    data: [],
    sources: Object.fromEntries(SOURCE_DEFINITIONS.map(def => [def.key, {
      key: def.key,
      label: def.label,
      status: 'idle',
      fileName: '',
      records: [],
      error: ''
    }]))
  };

  let recordSequence = 0;
  const VALIDATED_RECORDS = VALIDATED_RECORD_DEFINITIONS.map(data => createRecord('validated', data));
  state.data = VALIDATED_RECORDS.slice();

  function canAccess(){
    const role = window.CURRENT_USER && CURRENT_USER.role;
    return role === 'gerencia' || role === 'gerencia_director';
  }

  function normalizeText(value){
    return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  }

  function normalizeKey(value){
    return normalizeText(value)
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  function normalizeFileName(value){
    return normalizeKey(String(value || '').replace(/\.(xlsx|xls)$/i, ''));
  }

  function escapeHtml(value){
    return String(value == null ? '' : value)
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#39;');
  }

  function escapeAttr(value){
    return escapeHtml(value).replace(/`/g,'&#96;');
  }

  function toNumber(value){
    if(typeof value === 'number') return Number.isFinite(value) ? value : 0;
    if(value == null || value === '') return 0;
    let text = String(value)
      .replace(/\u00a0/g, ' ')
      .replace(/\(([^)]+)\)/, '-$1')
      .replace(/[^0-9,.-]/g, '')
      .trim();
    if(!text || text === '-' || text === '.' || text === ',') return 0;
    const comma = text.lastIndexOf(',');
    const dot = text.lastIndexOf('.');
    if(comma >= 0 && dot >= 0) {
      if(comma > dot) text = text.replace(/\./g,'').replace(',','.');
      else text = text.replace(/,/g,'');
    } else if(comma >= 0) {
      const decimals = text.length - comma - 1;
      text = decimals === 2 ? text.replace(/\./g,'').replace(',','.') : text.replace(/,/g,'');
    } else if(dot >= 0) {
      const decimals = text.length - dot - 1;
      if(decimals !== 2) text = text.replace(/\./g,'');
    }
    const number = Number(text);
    return Number.isFinite(number) ? number : 0;
  }

  function normalizePeriod(value){
    const text = normalizeText(value).toUpperCase().replace(/_/g,'-');
    if(!text) return 'Sin periodo';
    const compact = text.replace(/\s+/g,'');
    const match = compact.match(/FY(\d{2}|\d{4})-?Q([1-4])/);
    if(match) {
      const year = match[1].length === 4 ? match[1].slice(-2) : match[1];
      return `FY${year}-Q${match[2]}`;
    }
    const reverse = compact.match(/Q([1-4])-?FY(\d{2}|\d{4})/);
    if(reverse) {
      const year = reverse[2].length === 4 ? reverse[2].slice(-2) : reverse[2];
      return `FY${year}-Q${reverse[1]}`;
    }
    return text === 'SIN PERIODO' ? 'Sin periodo' : text;
  }

  function periodFromParts(yearValue, quarterValue){
    const yearMatch = normalizeText(yearValue).toUpperCase().match(/(?:FY)?(\d{2}|\d{4})/);
    const quarterMatch = normalizeText(quarterValue).toUpperCase().match(/Q?([1-4])/);
    if(!yearMatch || !quarterMatch) return 'Sin periodo';
    const year = yearMatch[1].length === 4 ? yearMatch[1].slice(-2) : yearMatch[1];
    return `FY${year}-Q${quarterMatch[1]}`;
  }

  function parseDate(value){
    if(value instanceof Date && !isNaN(value)) return value;
    if(typeof value === 'number' && value > 20000) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      return new Date(epoch.getTime() + value * 86400000);
    }
    const text = normalizeText(value);
    if(!text) return null;
    const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if(iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    const local = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
    if(local) {
      const year = Number(local[3]) < 100 ? 2000 + Number(local[3]) : Number(local[3]);
      return new Date(year, Number(local[2]) - 1, Number(local[1]));
    }
    const parsed = new Date(text);
    return isNaN(parsed) ? null : parsed;
  }

  function periodFromDate(value){
    const date = parseDate(value);
    if(!date) return 'Sin periodo';
    return `FY${String(date.getFullYear()).slice(-2)}-Q${Math.floor(date.getMonth() / 3) + 1}`;
  }

  function lenovoPeriod(program, quarter){
    const text = normalizeText(program).toUpperCase().replace(/\s+/g,'');
    const yearMatch = text.match(/FY(\d{2}|\d{4})/);
    const quarterMatch = normalizeText(quarter).toUpperCase().match(/Q([1-4])/) || text.match(/([1-4])Q/);
    if(!yearMatch || !quarterMatch) return 'Sin periodo';
    const rawYear = yearMatch[1];
    const year = rawYear.length === 4 ? rawYear.slice(-2) : rawYear;
    return `FY${year}-Q${quarterMatch[1]}`;
  }

  function periodSortValue(value){
    const match = normalizeText(value).match(/^FY(\d{2})-Q([1-4])$/i);
    if(!match) return -1;
    return Number(match[1]) * 10 + Number(match[2]);
  }

  function findSheetName(workbook, predicate){
    return (workbook.SheetNames || []).find(name => predicate(normalizeKey(name), name)) || '';
  }

  function sheetRows(workbook, sheetName){
    if(!sheetName || !workbook.Sheets[sheetName]) return [];
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: null,
      raw: true,
      blankrows: false
    });
  }

  function findHeaderIndex(rows, requiredKeys){
    const required = (requiredKeys || []).map(normalizeKey);
    const limit = Math.min(rows.length, 30);
    for(let index = 0; index < limit; index++) {
      const headers = (rows[index] || []).map(normalizeKey);
      if(required.every(key => headers.includes(key))) return index;
    }
    return -1;
  }

  function headerIndexes(row){
    const indexes = {};
    (row || []).forEach((value, index) => {
      const key = normalizeKey(value);
      if(!key) return;
      if(!indexes[key]) indexes[key] = [];
      indexes[key].push(index);
    });
    return indexes;
  }

  function valueByHeader(row, indexes, aliases, occurrence){
    for(const alias of aliases || []) {
      const positions = indexes[normalizeKey(alias)] || [];
      if(positions.length) {
        const position = positions[Math.min(occurrence || 0, positions.length - 1)];
        return row[position];
      }
    }
    return null;
  }

  function cleanState(value, fallback){
    const text = normalizeText(value);
    if(!text || /^\(en blanco\)$/i.test(text)) return fallback || 'Sin estado';
    return text;
  }

  function canonicalClaimState(value){
    const text = cleanState(value, 'Sin estado');
    const key = normalizeKey(text);
    if(key === 'allocated') return 'Asignado';
    if(key === 'expired') return 'Expirado';
    if(key === 'paid') return 'Pagado';
    return text;
  }

  function createRecord(sourceKey, data){
    const record = {
      id: `${sourceKey}-${++recordSequence}`,
      canal: cleanState(data.canal, 'Sin canal'),
      tipo: data.tipo === 'Rebate' ? 'Rebate' : 'Punto',
      programa: cleanState(data.programa, 'Sin programa'),
      periodo: normalizePeriod(data.periodo),
      valor: toNumber(data.valor),
      unidad: data.unidad === 'COP' ? 'COP' : data.unidad === 'USD' ? 'USD' : 'Puntos',
      estado: cleanState(data.estado, 'Sin estado'),
      clienteRef: cleanState(data.clienteRef, 'Sin referencia'),
      fuente: sourceKey,
      hoja: cleanState(data.hoja, ''),
      grupo: normalizeText(data.grupo || sourceKey),
      metrica: data.metrica === 'balance' ? 'balance' : 'movement'
    };
    return record;
  }

  function parseMyRewards(workbook){
    const records = [];
    const unclaimedName = findSheetName(workbook, key => key === 'unclaimed sales');
    const unclaimedRows = sheetRows(workbook, unclaimedName);
    const unclaimedHeader = findHeaderIndex(unclaimedRows, ['Quarter','Promotion Name','Unclaimed Points']);
    if(unclaimedHeader >= 0) {
      const indexes = headerIndexes(unclaimedRows[unclaimedHeader]);
      for(let rowIndex = unclaimedHeader + 1; rowIndex < unclaimedRows.length; rowIndex++) {
        const row = unclaimedRows[rowIndex] || [];
        const rawUnclaimed = valueByHeader(row, indexes, ['Unclaimed Points']);
        const value = rawUnclaimed == null
          ? toNumber(valueByHeader(row, indexes, ['Points']))
          : toNumber(rawUnclaimed);
        if(value <= 0) continue;
        records.push(createRecord('myrewards', {
          canal: 'Dell',
          tipo: 'Punto',
          metrica: 'balance',
          programa: valueByHeader(row, indexes, ['Promotion Name']),
          periodo: valueByHeader(row, indexes, ['Quarter']),
          valor: value,
          unidad: 'Puntos',
          estado: 'Disponible',
          clienteRef: valueByHeader(row, indexes, ['Customer Name','Order Number','SNS Transaction Id']),
          hoja: unclaimedName
        }));
      }
    }

    const claimsName = findSheetName(workbook, key => key === 'claims summary');
    const claimRows = sheetRows(workbook, claimsName);
    const claimHeader = findHeaderIndex(claimRows, ['Quarter','Promotion Name','MyR Points','Claim Status']);
    if(claimHeader >= 0) {
      const indexes = headerIndexes(claimRows[claimHeader]);
      for(let rowIndex = claimHeader + 1; rowIndex < claimRows.length; rowIndex++) {
        const row = claimRows[rowIndex] || [];
        const value = toNumber(valueByHeader(row, indexes, ['MyR Points']));
        if(value <= 0) continue;
        const claimant = valueByHeader(row, indexes, ['Claimant']);
        const claimId = valueByHeader(row, indexes, ['Claim ID']);
        const customer = valueByHeader(row, indexes, ['Customer Name']);
        records.push(createRecord('myrewards', {
          canal: 'Dell',
          tipo: 'Punto',
          metrica: normalizeKey(valueByHeader(row, indexes, ['Claim Status'])).includes('unclaimed') ? 'balance' : 'movement',
          programa: valueByHeader(row, indexes, ['Promotion Name']),
          periodo: valueByHeader(row, indexes, ['Quarter']),
          valor: value,
          unidad: 'Puntos',
          estado: canonicalClaimState(valueByHeader(row, indexes, ['Claim Status'])),
          clienteRef: [customer || claimant, claimId ? `Claim ${claimId}` : ''].filter(Boolean).join(' · '),
          hoja: claimsName
        }));
      }
    }
    return records;
  }

  function parseDellClaims(workbook){
    const records = [];
    const sheetName = findSheetName(workbook, key => key === 'table1');
    const rows = sheetRows(workbook, sheetName);
    const header = findHeaderIndex(rows, ['Fund Name','Claim Status Name','Approved to Pay']);
    if(header < 0) return records;
    const indexes = headerIndexes(rows[header]);
    for(let rowIndex = header + 1; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex] || [];
      const program = valueByHeader(row, indexes, ['Fund Name']);
      if(!normalizeText(program)) continue;
      const payment = cleanState(valueByHeader(row, indexes, ['Claim Payment Status']), '');
      const claim = cleanState(valueByHeader(row, indexes, ['Claim Status Name']), '');
      const statusParts = [...new Set([payment, claim].filter(Boolean))];
      const currencyRaw = normalizeText(valueByHeader(row, indexes, ['Budget Currency','Claimed Currency'])).toUpperCase();
      const unit = currencyRaw.includes('COP') ? 'COP' : 'USD';
      const period = periodFromParts(
        valueByHeader(row, indexes, ['Fund Year','Execution Year']),
        valueByHeader(row, indexes, ['Fund Quarter','Execution Quarter'])
      );
      const claimId = valueByHeader(row, indexes, ['Claim ID']);
      const reference = valueByHeader(row, indexes, ['Reference','Consumer Reference Number']);
      records.push(createRecord('claims', {
        canal: 'Dell',
        tipo: 'Rebate',
        programa: program,
        periodo: period,
        valor: valueByHeader(row, indexes, ['Approved to Pay']),
        unidad: unit,
        estado: statusParts.join(' / ') || 'Sin estado',
        clienteRef: [claimId ? `Claim ${claimId}` : '', reference].filter(Boolean).join(' · ') || valueByHeader(row, indexes, ['Partner Name']),
        hoja: sheetName
      }));
    }
    return records;
  }

  function parseLenovo(workbook){
    const records = [];
    const sheetName = findSheetName(workbook, key => key === 'program list');
    const rows = sheetRows(workbook, sheetName);
    const header = findHeaderIndex(rows, ['Program Name','Lenovo Quarter','Rebate Earned','Estado']);
    if(header < 0) return records;
    const indexes = headerIndexes(rows[header]);
    for(let rowIndex = header + 1; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex] || [];
      const program = valueByHeader(row, indexes, ['Program Name']);
      if(!normalizeText(program)) continue;
      const reference = valueByHeader(row, indexes, ['Número de referencia','Numero de referencia']);
      const creditNote = valueByHeader(row, indexes, ['Número de nota de crédito','Numero de nota de credito']);
      const bankRef = valueByHeader(row, indexes, ['Número de referencia del banco','Numero de referencia del banco']);
      records.push(createRecord('lenovo', {
        canal: 'Lenovo',
        tipo: 'Rebate',
        programa: program,
        periodo: lenovoPeriod(program, valueByHeader(row, indexes, ['Lenovo Quarter'])),
        valor: valueByHeader(row, indexes, ['Rebate Earned']),
        unidad: 'USD',
        estado: valueByHeader(row, indexes, ['Estado']),
        clienteRef: creditNote || bankRef || reference,
        hoja: sheetName
      }));
    }
    return records;
  }

  function parsePlatforms(workbook){
    const records = [];

    const epsonName = findSheetName(workbook, key => key === 'epson');
    const epsonRows = sheetRows(workbook, epsonName);
    const epsonHeader = findHeaderIndex(epsonRows, ['ID','Programa','Valor de Rebate']);
    if(epsonHeader >= 0) {
      for(let rowIndex = epsonHeader + 1; rowIndex < epsonRows.length; rowIndex++) {
        const row = epsonRows[rowIndex] || [];
        const id = normalizeText(row[0]);
        const program = normalizeText(row[3]);
        if(!id || !program) continue;
        const stateValue = cleanState(row[9], cleanState(row[5], 'Generado'));
        const reference = [row[6], row[8] ? `NC ${normalizeText(row[8])}` : '', `ID ${id}`].filter(Boolean).join(' · ');
        const period = periodFromDate(row[1] || row[2]);
        records.push(createRecord('platforms', {
          canal: 'Epson',
          tipo: 'Rebate',
          programa: program,
          periodo: period,
          valor: row[4],
          unidad: 'USD',
          estado: stateValue,
          clienteRef: reference,
          hoja: epsonName
        }));
        const copValue = toNumber(row[7]) || toNumber(row[10]);
        if(copValue > 0) {
          records.push(createRecord('platforms', {
            canal: 'Epson',
            tipo: 'Rebate',
            programa: program,
            periodo: period,
            valor: copValue,
            unidad: 'COP',
            estado: stateValue,
            clienteRef: reference,
            hoja: epsonName
          }));
        }
      }
    }

    const hpeInstantName = findSheetName(workbook, key => key === 'hpe instan on');
    const hpeInstantRows = sheetRows(workbook, hpeInstantName);
    const hpeHeader = findHeaderIndex(hpeInstantRows, ['Fecha','Descripción','Incentivos redimidos']);
    if(hpeHeader >= 0) {
      const indexes = headerIndexes(hpeInstantRows[hpeHeader]);
      for(let rowIndex = hpeHeader + 1; rowIndex < hpeInstantRows.length; rowIndex++) {
        const row = hpeInstantRows[rowIndex] || [];
        const program = valueByHeader(row, indexes, ['Descripción','Descripcion']);
        const value = toNumber(valueByHeader(row, indexes, ['Incentivos redimidos']));
        if(!normalizeText(program) || value <= 0) continue;
        const extractedPeriod = normalizeText(program).match(/FY\d{2}Q[1-4]/i);
        records.push(createRecord('platforms', {
          canal: 'HPE',
          tipo: 'Punto',
          programa: program,
          periodo: extractedPeriod ? extractedPeriod[0] : periodFromDate(valueByHeader(row, indexes, ['Fecha'])),
          valor: value,
          unidad: 'Puntos',
          estado: 'Redimido',
          clienteRef: valueByHeader(row, indexes, ['Identificación de canje','Identificacion de canje']),
          hoja: hpeInstantName
        }));
      }
    }

    const hpeName = findSheetName(workbook, key => key === 'hpe');
    const hpeRows = sheetRows(workbook, hpeName);
    if(hpeRows.length > 1) {
      for(let rowIndex = 1; rowIndex < hpeRows.length; rowIndex++) {
        const row = hpeRows[rowIndex] || [];
        const concept = normalizeText(row[0]);
        if(!concept) continue;
        const period = periodFromDate(row[2]);
        const redeemed = toNumber(row[1]);
        const available = toNumber(row[4]);
        if(redeemed > 0) {
          records.push(createRecord('platforms', {
            canal: 'HPE', tipo: 'Punto', programa: concept, periodo: period,
            valor: redeemed, unidad: 'Puntos', estado: 'Redimido',
            clienteRef: normalizeText(row[2]) || 'Plataforma HPE', hoja: hpeName
          }));
        }
        if(available > 0) {
          records.push(createRecord('platforms', {
            canal: 'HPE', tipo: 'Punto', programa: 'Saldo disponible HPE', periodo: period,
            metrica: 'balance', valor: available, unidad: 'Puntos', estado: 'Disponible',
            clienteRef: concept, hoja: hpeName
          }));
        }
      }
    }

    const intelName = findSheetName(workbook, key => key === 'intel');
    const intelRows = sheetRows(workbook, intelName);
    const intelState = intelRows.flat().map(normalizeText).find(value => /pendiente/i.test(value)) || 'Pendiente';
    for(let rowIndex = 1; rowIndex < intelRows.length; rowIndex++) {
      const row = intelRows[rowIndex] || [];
      const quantity = toNumber(row[0]);
      const denomination = toNumber(row[1]);
      const value = toNumber(row[2]);
      if(quantity <= 0 || value <= 0) continue;
      records.push(createRecord('platforms', {
        canal: 'Intel',
        tipo: 'Punto',
        metrica: 'balance',
        programa: `Incentivo Intel de ${formatInteger(denomination)} puntos`,
        periodo: 'Sin periodo',
        valor: value,
        unidad: 'Puntos',
        estado: intelState,
        clienteRef: `Cantidad: ${formatInteger(quantity)}`,
        hoja: intelName
      }));
    }

    return records;
  }

  function parseWorkbook(sourceKey, workbook){
    if(sourceKey === 'myrewards') return parseMyRewards(workbook);
    if(sourceKey === 'claims') return parseDellClaims(workbook);
    if(sourceKey === 'lenovo') return parseLenovo(workbook);
    if(sourceKey === 'platforms') return parsePlatforms(workbook);
    return [];
  }

  function recognizeSource(fileName){
    const normalized = normalizeFileName(fileName);
    return SOURCE_DEFINITIONS.find(def => def.matches(normalized)) || null;
  }

  function resetSources(){
    SOURCE_DEFINITIONS.forEach(def => {
      state.sources[def.key] = {
        key: def.key,
        label: def.label,
        status: 'loading',
        fileName: '',
        records: [],
        error: ''
      };
    });
    state.data = VALIDATED_RECORDS.slice();
    recordSequence = 0;
  }

  function rebuildData(){
    const loadedGroups = new Set(SOURCE_DEFINITIONS
      .filter(def => state.sources[def.key].status === 'loaded' && state.sources[def.key].records.length)
      .map(def => def.key));
    const validated = VALIDATED_RECORDS.filter(row => !loadedGroups.has(row.grupo));
    const live = SOURCE_DEFINITIONS.flatMap(def => state.sources[def.key].records || []);
    state.data = [...validated, ...live];
    state.page = 1;
  }

  function setSourceResult(definition, fileName, records, error){
    state.sources[definition.key] = {
      key: definition.key,
      label: definition.label,
      status: error ? 'error' : fileName ? 'loaded' : 'missing',
      fileName: fileName || '',
      records: records || [],
      error: error || ''
    };
  }

  function sortDriveItems(items){
    return [...items].sort((a,b) => {
      const aDate = new Date(a.lastModifiedDateTime || 0).getTime();
      const bDate = new Date(b.lastModifiedDateTime || 0).getTime();
      return bDate - aDate || String(b.name || '').localeCompare(String(a.name || ''));
    });
  }

  async function findSourceItemsInFabricantes(siteId, token){
    const forecastBasePath = await getForecastBasePath(siteId, token);
    const folderPath = joinGraphPath(forecastBasePath, SHAREPOINT_FOLDER_NAME);
    const response = await fetch(buildGraphRootUrl(siteId, folderPath, 'children?$top=200'), {
      headers: { Authorization: 'Bearer ' + token }
    });
    const payload = await response.json().catch(() => ({}));
    if(!response.ok) {
      throw new Error((payload.error && payload.error.message) || `No se pudo consultar ${folderPath}.`);
    }

    const files = (payload.value || []).filter(item =>
      item && item.name && !item.name.startsWith('~$') && /\.xlsx?$/i.test(item.name)
    );
    const itemsBySource = new Map();
    SOURCE_DEFINITIONS.forEach(definition => {
      const matches = files.filter(item => definition.matches(normalizeFileName(item.name)));
      itemsBySource.set(definition.key, sortDriveItems(matches)[0] || null);
    });
    console.log('[PROGRAMAS] carpeta SharePoint', folderPath, files.length, 'archivos Excel');
    return itemsBySource;
  }

  async function findSourceItem(definition, drives, token){
    const found = new Map();
    for(const drive of drives) {
      for(const query of definition.queries) {
        try {
          const response = await fetch(buildDriveRootSearchUrl(drive.driveId, query), {
            headers: { Authorization: 'Bearer ' + token }
          });
          const payload = await response.json().catch(() => ({}));
          if(!response.ok) {
            console.warn('[PROGRAMAS SEARCH]', definition.key, drive.driveName, payload);
            continue;
          }
          (payload.value || []).forEach(item => {
            if(!item || !item.name || item.name.startsWith('~$') || !/\.xlsx?$/i.test(item.name)) return;
            if(!definition.matches(normalizeFileName(item.name))) return;
            const enriched = {
              ...item,
              parentReference: {
                ...(item.parentReference || {}),
                driveId: item.parentReference && item.parentReference.driveId || drive.driveId
              },
              __SEARCH_DRIVE_ID: drive.driveId,
              __SEARCH_DRIVE_NAME: drive.driveName
            };
            found.set(getDriveItemKey(enriched), enriched);
          });
        } catch(error) {
          console.warn('[PROGRAMAS SEARCH]', definition.key, error);
        }
      }
    }
    return sortDriveItems(Array.from(found.values()))[0] || null;
  }

  async function loadSourceItem(definition, item, token){
    if(!item) {
      setSourceResult(definition, '', [], '');
      return;
    }
    try {
      const file = await ensureDriveItemDownloadUrl(item, token);
      if(!file || !file['@microsoft.graph.downloadUrl']) throw new Error('El archivo no tiene URL de descarga.');
      const response = await fetch(file['@microsoft.graph.downloadUrl']);
      if(!response.ok) throw new Error(`No se pudo descargar (${response.status}).`);
      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, { type:'array', cellDates:true });
      const records = parseWorkbook(definition.key, workbook);
      setSourceResult(definition, file.name, records, '');
      console.log('[PROGRAMAS]', file.name, records.length, 'registros');
    } catch(error) {
      console.warn('[PROGRAMAS]', definition.key, error);
      setSourceResult(definition, item.name || '', [], error.message || String(error));
    }
  }

  async function loadFromSharePoint(siteId, token){
    if(!canAccess()) return;
    if(typeof XLSX === 'undefined') throw new Error('La libreria XLSX no esta disponible.');
    state.loading = true;
    state.error = '';
    resetSources();
    render();
    try {
      let folderItems = new Map();
      let drives = null;
      try {
        if(typeof updateLoadingStatus === 'function') updateLoadingStatus(`Consultando ${SHAREPOINT_FOLDER_NAME}...`);
        folderItems = await findSourceItemsInFabricantes(siteId, token);
      } catch(folderError) {
        console.warn('[PROGRAMAS] no se pudo consultar la carpeta Fabricantes; se usara la busqueda general', folderError);
      }
      for(const definition of SOURCE_DEFINITIONS) {
        let item = folderItems.get(definition.key) || null;
        if(!item) {
          if(typeof updateLoadingStatus === 'function') updateLoadingStatus(`Buscando respaldo: ${definition.label}...`);
          if(!drives) drives = await getSearchDriveCandidates(siteId, token);
          item = await findSourceItem(definition, drives, token);
        } else if(typeof updateLoadingStatus === 'function') {
          updateLoadingStatus(`Leyendo Fabricantes: ${definition.label}...`);
        }
        await loadSourceItem(definition, item, token);
        rebuildData();
        render();
      }
    } catch(error) {
      state.error = error.message || String(error);
      console.warn('[PROGRAMAS]', error);
    } finally {
      state.loading = false;
      rebuildData();
      render();
    }
  }

  async function reloadFromSharePoint(){
    if(!canAccess() || state.loading) return;
    state.loading = true;
    state.error = '';
    render();
    try {
      const siteId = await getSiteId();
      const token = await getToken(['Files.Read.All']);
      await loadFromSharePoint(siteId, token);
    } catch(error) {
      state.error = error.message || String(error);
      state.loading = false;
      render();
    }
  }

  function openLocalFiles(){
    if(!canAccess()) return;
    const input = document.getElementById('program-channel-file-input');
    if(input) input.click();
  }

  async function handleLocalFiles(fileList){
    if(!canAccess()) return;
    const files = Array.from(fileList || []);
    if(!files.length) return;
    state.loading = true;
    state.error = '';
    render();
    const recognized = [];
    for(const file of files) {
      const definition = recognizeSource(file.name);
      if(!definition) continue;
      recognized.push(definition.key);
      state.sources[definition.key].status = 'loading';
      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type:'array', cellDates:true });
        const records = parseWorkbook(definition.key, workbook);
        setSourceResult(definition, file.name, records, '');
      } catch(error) {
        setSourceResult(definition, file.name, [], error.message || String(error));
      }
    }
    if(!recognized.length) state.error = 'Ningun archivo coincide con los cuatro reportes esperados.';
    state.loading = false;
    rebuildData();
    const input = document.getElementById('program-channel-file-input');
    if(input) input.value = '';
    render();
  }

  function setMode(mode){
    if(mode !== 'Punto' && mode !== 'Rebate') return;
    state.mode = mode;
    state.filters = {
      canal: '',
      periodo: '',
      estado: '',
      unidad: mode === 'Punto' ? 'Puntos' : 'USD'
    };
    state.page = 1;
    render();
  }

  function setFilter(key, value){
    if(!Object.prototype.hasOwnProperty.call(state.filters, key)) return;
    state.filters[key] = normalizeText(value);
    state.page = 1;
    render();
  }

  function clearFilters(){
    state.filters.canal = '';
    state.filters.periodo = '';
    state.filters.estado = '';
    state.filters.unidad = state.mode === 'Punto' ? 'Puntos' : 'USD';
    state.page = 1;
    render();
  }

  function setPage(page){
    state.page = Math.max(1, Number(page) || 1);
    renderTable(getVisibleRows());
  }

  function typeRows(){
    return state.data.filter(row => row.tipo === state.mode);
  }

  function rowsWithoutUnit(){
    return typeRows().filter(row =>
      (!state.filters.canal || row.canal === state.filters.canal) &&
      (!state.filters.periodo || row.periodo === state.filters.periodo) &&
      (!state.filters.estado || row.estado === state.filters.estado)
    );
  }

  function getVisibleRows(){
    const rows = rowsWithoutUnit().filter(row =>
      state.mode === 'Punto' || !state.filters.unidad || row.unidad === state.filters.unidad
    );
    return rows.sort((a,b) =>
      periodSortValue(b.periodo) - periodSortValue(a.periodo) ||
      b.valor - a.valor ||
      a.canal.localeCompare(b.canal)
    );
  }

  function uniqueSorted(values, sorter){
    const list = [...new Set(values.map(normalizeText).filter(Boolean))];
    return sorter ? list.sort(sorter) : list.sort((a,b) => a.localeCompare(b, 'es'));
  }

  function syncSelect(id, options, selected, emptyLabel){
    const select = document.getElementById(id);
    if(!select) return '';
    const available = options.includes(selected) ? selected : '';
    select.innerHTML = `<option value="">${escapeHtml(emptyLabel)}</option>` + options.map(option =>
      `<option value="${escapeAttr(option)}"${option === available ? ' selected' : ''}>${escapeHtml(option)}</option>`
    ).join('');
    select.value = available;
    return available;
  }

  function renderChannelChips(rows){
    const host = document.getElementById('program-channel-channel-chips');
    const context = document.getElementById('program-channel-channel-context');
    if(!host) return;
    const discovered = uniqueSorted((rows || []).map(row => row.canal));
    const channels = [...CHANNELS, ...discovered.filter(channel => !CHANNELS.includes(channel))];
    if(state.filters.canal && !channels.includes(state.filters.canal)) state.filters.canal = '';
    host.innerHTML = [{ value:'', label:'Todos' }, ...channels.map(channel => ({ value:channel, label:channel }))]
      .map(item => {
        const active = state.filters.canal === item.value;
        return `<button type="button" class="program-channel-channel-btn${active ? ' active' : ''}" aria-pressed="${active ? 'true' : 'false'}" onclick="ProgramChannelModule.setFilter('canal',${JSON.stringify(item.value)})">${escapeHtml(item.label)}</button>`;
      }).join('');
    if(context) {
      const count = state.filters.canal
        ? rows.filter(row => row.canal === state.filters.canal).length
        : rows.length;
      context.style.display = 'inline-flex';
      context.textContent = `${count.toLocaleString('es-CO')} registros`;
    }
  }

  function syncFilters(){
    const rows = typeRows();
    renderChannelChips(rows);
    const channelRows = state.filters.canal ? rows.filter(row => row.canal === state.filters.canal) : rows;
    state.filters.periodo = syncSelect(
      'program-channel-filter-period',
      uniqueSorted(channelRows.map(row => row.periodo), (a,b) => periodSortValue(b) - periodSortValue(a) || a.localeCompare(b)),
      state.filters.periodo,
      'Todos los periodos'
    );
    const periodRows = state.filters.periodo ? channelRows.filter(row => row.periodo === state.filters.periodo) : channelRows;
    state.filters.estado = syncSelect(
      'program-channel-filter-status',
      uniqueSorted(periodRows.map(row => row.estado)),
      state.filters.estado,
      'Todos los estados'
    );

    const unitWrap = document.getElementById('program-channel-unit-filter');
    const unitSelect = document.getElementById('program-channel-filter-unit');
    if(unitWrap) unitWrap.style.display = '';
    const allowedUnits = state.mode === 'Punto' ? ['Puntos','COP'] : ['USD','COP'];
    const filteredForUnits = periodRows.filter(row => !state.filters.estado || row.estado === state.filters.estado);
    let units = uniqueSorted(filteredForUnits.map(row => row.unidad)).filter(unit => allowedUnits.includes(unit));
    if(!units.length) units = uniqueSorted(rows.map(row => row.unidad)).filter(unit => allowedUnits.includes(unit));
    const preferredUnit = state.mode === 'Punto' ? 'Puntos' : 'USD';
    if(!units.includes(state.filters.unidad)) state.filters.unidad = units.includes(preferredUnit) ? preferredUnit : (units[0] || preferredUnit);
    if(unitSelect) {
      unitSelect.innerHTML = units.map(unit => `<option value="${unit}"${unit === state.filters.unidad ? ' selected' : ''}>${unit}</option>`).join('');
      unitSelect.value = state.filters.unidad;
    }

    const typeLabel = document.getElementById('program-channel-filter-type');
    if(typeLabel) typeLabel.textContent = state.mode;
  }

  function statusGroup(value){
    const key = normalizeKey(value);
    if(/expir/.test(key)) return 'expired';
    if(/declin|cancel|rechaz|perdid/.test(key)) return 'rejected';
    if(/pend|review|released|dispon|unclaimed|generad|vigent|oportunidad|alerta/.test(key)) return 'pending';
    if(/redimid|asignad|allocated|cargado/.test(key)) return 'redeemed';
    if(/pagad|paid|procesad|complete/.test(key)) return 'paid';
    return 'other';
  }

  function formatInteger(value){
    return Math.round(toNumber(value)).toLocaleString('es-CO');
  }

  function formatMoney(value, unit){
    const number = toNumber(value);
    if(unit === 'COP') return '$ ' + Math.round(number).toLocaleString('es-CO');
    if(unit === 'USD') return 'USD ' + number.toLocaleString('es-CO', { minimumFractionDigits:2, maximumFractionDigits:2 });
    return formatInteger(number);
  }

  function formatCompact(value, unit){
    const number = toNumber(value);
    const abs = Math.abs(number);
    let body = '';
    if(abs >= 1e9) body = (number / 1e9).toFixed(2) + ' B';
    else if(abs >= 1e6) body = (number / 1e6).toFixed(2) + ' M';
    else if(abs >= 1e3) body = (number / 1e3).toFixed(abs >= 1e5 ? 0 : 1) + ' K';
    else body = number.toLocaleString('es-CO', { maximumFractionDigits:2 });
    if(unit === 'COP') return '$ ' + body;
    if(unit === 'USD') return 'USD ' + body;
    return body;
  }

  function sumRows(rows){
    return (rows || []).reduce((sum, row) => sum + toNumber(row.valor), 0);
  }

  function kpiCard(label, value, subtext, accent, extraClass){
    return `<article class="kpi program-channel-kpi ${escapeAttr(extraClass || '')}" style="--program-accent:${accent}">
      <div class="kpi-label">${escapeHtml(label)}</div>
      <div class="kpi-val program-channel-kpi-value" title="${escapeAttr(value)}">${escapeHtml(value)}</div>
      <div class="kpi-sub">${escapeHtml(subtext)}</div>
    </article>`;
  }

  function renderKpis(){
    const host = document.getElementById('program-channel-kpis');
    if(!host) return;
    const allUnitsRows = rowsWithoutUnit();
    const hasDetailFilters = Boolean(state.filters.periodo || state.filters.estado);

    if(state.mode === 'Punto' && state.filters.unidad === 'COP') {
      const copRows = allUnitsRows.filter(row => row.unidad === 'COP');
      const pendingCop = sumRows(copRows.filter(row => statusGroup(row.estado) === 'pending'));
      host.innerHTML = [
        kpiCard('Incentivos COP', formatMoney(sumRows(copRows), 'COP'), 'Separados de los puntos', '#8B5FC8'),
        kpiCard('Vigentes / pendientes COP', formatMoney(pendingCop, 'COP'), `${copRows.length} programa(s)`, '#F0A020'),
        kpiCard('Bono viajero Lenovo', formatMoney(BUSINESS_METRICS.sed.travelerBonusCop, 'COP'), 'Lenovo Adventure', '#2D4FD6'),
        kpiCard('Vigencia bonos SED', `${BUSINESS_METRICS.sed.validityMonths} meses`, 'Seguimiento requerido', '#E84040')
      ].join('');
      return;
    }

    if(state.mode === 'Punto') {
      const pointRows = allUnitsRows.filter(row => row.unidad === 'Puntos');
      const movements = pointRows.filter(row => row.metrica !== 'balance');
      const balances = pointRows.filter(row => row.metrica === 'balance');
      const pending = sumRows(pointRows.filter(row => statusGroup(row.estado) === 'pending'));
      const redeemed = sumRows(pointRows.filter(row => /redimid|paid|procesad|cargado/i.test(normalizeKey(row.estado))));
      const expired = sumRows(pointRows.filter(row => statusGroup(row.estado) === 'expired'));

      if(!hasDetailFilters && state.filters.canal === 'Dell') {
        host.innerHTML = [
          kpiCard('Puntos históricos', formatInteger(sumRows(movements)), 'FY22-Q1 a FY23-Q2', '#2ABFDF'),
          kpiCard('Pendientes por redimir', formatInteger(BUSINESS_METRICS.dell.unredeemedPoints), 'Acción requerida', '#F0A020', 'program-channel-kpi-alert'),
          kpiCard('Estado clave', 'Unclaimed', 'Dell MyRewards', '#E84040'),
          kpiCard('Unidad', 'Puntos', 'Nunca se suman con dinero', '#8B5FC8')
        ].join('');
        return;
      }
      if(!hasDetailFilters && state.filters.canal === 'HPE') {
        host.innerHTML = [
          kpiCard('HPE Instant On', formatInteger(BUSINESS_METRICS.hpe.availablePoints), 'Puntos disponibles', '#2ABFDF'),
          kpiCard('Visa Prepaid Card', `USD ${BUSINESS_METRICS.hpe.visaUsd}`, 'Equivalencia informativa, no sumada', '#8B5FC8'),
          kpiCard('Cash bonus FY26-Q2', '242', 'Puntos del trimestre', '#0DBF82'),
          kpiCard('Mejor trimestre', '403', 'FY25-Q3', '#F0A020')
        ].join('');
        return;
      }
      if(!hasDetailFilters && state.filters.canal === 'Lenovo') {
        host.innerHTML = [
          kpiCard('Plan Ultra Aguinaldo WS', '37', 'Puntos disponibles', '#2ABFDF'),
          kpiCard('Bono viajero', formatMoney(BUSINESS_METRICS.sed.travelerBonusCop, 'COP'), 'Valor COP separado', '#8B5FC8'),
          kpiCard('Vigencia', `${BUSINESS_METRICS.sed.validityMonths} meses`, 'Bonos SED', '#F0A020'),
          kpiCard('Oportunidad', 'Workstation', 'Acelerar activación comercial', '#0DBF82')
        ].join('');
        return;
      }

      host.innerHTML = [
        kpiCard('Puntos acumulados', formatInteger(sumRows(movements)), `${movements.length} movimientos`, '#2ABFDF'),
        kpiCard('Redimidos / entregados', formatInteger(redeemed), 'No incluye asignados', '#0DBF82'),
        kpiCard('Disponibles / pendientes', formatInteger(pending || sumRows(balances)), 'Sin redimir', '#F0A020'),
        state.filters.canal
          ? kpiCard('Expirados', formatInteger(expired), 'Puntos vencidos', '#E84040')
          : kpiCard('Dell por redimir', formatInteger(BUSINESS_METRICS.dell.unredeemedPoints), 'Acción requerida', '#E84040', 'program-channel-kpi-alert')
      ].join('');
      return;
    }

    if(!hasDetailFilters && state.filters.unidad === 'USD' && !state.filters.canal) {
      host.innerHTML = [
        kpiCard('Rebates Lenovo 2026', formatMoney(BUSINESS_METRICS.lenovo.actual2026, 'USD'), 'Actual', '#8B5FC8'),
        kpiCard('Proyección de cierre', formatMoney(BUSINESS_METRICS.lenovo.projection2026, 'USD'), `+${BUSINESS_METRICS.lenovo.yoy2026}% YoY`, '#0DBF82'),
        kpiCard('Dell aprobado a pagar', formatMoney(BUSINESS_METRICS.dell.approved, 'USD'), 'FY26–FY27', '#2D4FD6'),
        kpiCard('ASUS cumplimiento', `${BUSINESS_METRICS.asus.compliance}%`, `Alerta · meta ${formatMoney(BUSINESS_METRICS.asus.quota, 'USD')}`, '#E84040', 'program-channel-kpi-alert')
      ].join('');
      return;
    }

    if(!hasDetailFilters && state.filters.unidad === 'USD' && state.filters.canal === 'Lenovo') {
      host.innerHTML = [
        kpiCard('Actual 2026', formatMoney(BUSINESS_METRICS.lenovo.actual2026, 'USD'), 'Lenovo 360 Engage', '#2ABFDF'),
        kpiCard('Proyección de cierre', formatMoney(BUSINESS_METRICS.lenovo.projection2026, 'USD'), 'Mejor año proyectado', '#0DBF82'),
        kpiCard('Crecimiento YoY', `+${BUSINESS_METRICS.lenovo.yoy2026}%`, 'Vs. cierre 2025', '#8B5FC8'),
        kpiCard('Oportunidad', 'WKS Expert', 'Workstation subexplotado', '#F0A020')
      ].join('');
      return;
    }

    if(!hasDetailFilters && state.filters.unidad === 'USD' && state.filters.canal === 'Dell') {
      const dellRows = allUnitsRows.filter(row => row.unidad === 'USD');
      const paid = sumRows(dellRows.filter(row => ['paid','redeemed'].includes(statusGroup(row.estado))));
      const openFunds = dellRows.filter(row => ['pending'].includes(statusGroup(row.estado)) && row.valor === 0).length;
      host.innerHTML = [
        kpiCard('Aprobado a pagar', formatMoney(BUSINESS_METRICS.dell.approved, 'USD'), 'Total validado', '#2ABFDF'),
        kpiCard('Paid / Complete', formatMoney(paid, 'USD'), 'FY26 ejecutado', '#0DBF82'),
        kpiCard('Fondos sin convertir', formatInteger(openFunds), 'Released / Pending Final Review', '#F0A020'),
        kpiCard('Declined', formatInteger(dellRows.filter(row => statusGroup(row.estado) === 'rejected').length), 'Requiere trazabilidad', '#E84040')
      ].join('');
      return;
    }

    if(!hasDetailFilters && state.filters.unidad === 'USD' && state.filters.canal === 'ASUS') {
      host.innerHTML = [
        kpiCard('Cuota 2026', formatMoney(BUSINESS_METRICS.asus.quota, 'USD'), 'AGP Silver', '#2D4FD6'),
        kpiCard('Ventas a 30 jun', formatMoney(BUSINESS_METRICS.asus.sales, 'USD'), '1 ene–30 jun 2026', '#2ABFDF'),
        kpiCard('Unidades facturadas', formatInteger(BUSINESS_METRICS.asus.units), 'Portafolio elegible', '#8B5FC8'),
        kpiCard('Cumplimiento', `${BUSINESS_METRICS.asus.compliance}%`, 'Alerta', '#E84040', 'program-channel-kpi-alert')
      ].join('');
      return;
    }

    const selectedUnitRows = allUnitsRows.filter(row => row.unidad === state.filters.unidad);
    const paid = sumRows(selectedUnitRows.filter(row => ['paid','redeemed'].includes(statusGroup(row.estado))));
    const pending = sumRows(selectedUnitRows.filter(row => statusGroup(row.estado) === 'pending'));
    const rejected = selectedUnitRows.filter(row => ['expired','rejected'].includes(statusGroup(row.estado))).length;
    host.innerHTML = [
      kpiCard(`Total rebates ${state.filters.unidad}`, formatMoney(sumRows(selectedUnitRows), state.filters.unidad), `${selectedUnitRows.length} registros`, '#2ABFDF'),
      kpiCard(`Pagados / procesados ${state.filters.unidad}`, formatMoney(paid, state.filters.unidad), 'Estado completado', '#0DBF82'),
      kpiCard(`Pendientes ${state.filters.unidad}`, formatMoney(pending, state.filters.unidad), 'Por gestionar', '#F0A020'),
      kpiCard('Declined / Expired', formatInteger(rejected), 'Registros en alerta', '#E84040')
    ].join('');
  }

  function renderBarsChart(rows){
    const title = document.getElementById('program-channel-bars-title');
    const host = document.getElementById('program-channel-bars');
    const executiveOverview = state.mode === 'Rebate' && state.filters.unidad === 'USD' &&
      !state.filters.canal && !state.filters.periodo && !state.filters.estado;
    if(title) title.innerHTML = executiveOverview
      ? 'Rebate por canal <span>USD · proyección Lenovo / aprobado Dell</span>'
      : `${state.mode === 'Punto' ? 'Puntos' : 'Rebates'} por canal <span>${escapeHtml(state.filters.unidad)}</span>`;
    if(!host) return;
    const totals = new Map();
    const chartRows = state.mode === 'Punto' ? rows.filter(row => row.metrica !== 'balance') : rows;
    chartRows.forEach(row => totals.set(row.canal, (totals.get(row.canal) || 0) + row.valor));
    let items = [...totals.entries()].map(([name,val]) => ({name,val})).sort((a,b) => b.val - a.val);
    if(executiveOverview) {
      items = [
        { name:'Lenovo', val:BUSINESS_METRICS.lenovo.projection2026 },
        { name:'Dell', val:BUSINESS_METRICS.dell.approved },
        { name:'Epson', val:sumRows(rows.filter(row => row.canal === 'Epson')) },
        { name:'ASUS', val:0 }
      ];
    }
    if(!items.length) {
      host.innerHTML = '<div class="program-channel-empty">Sin datos para los filtros seleccionados.</div>';
      return;
    }
    if(typeof renderBars === 'function') {
      const palette = typeof COLORS !== 'undefined' ? COLORS : ['#2D4FD6','#8B5FC8','#2ABFDF','#0DBF82'];
      renderBars('program-channel-bars', items, palette, value =>
        state.mode === 'Punto' ? formatInteger(value) : formatCompact(value, state.filters.unidad), {
          nameClass: 'w100',
          getOnClick: item => `ProgramChannelModule.setFilter('canal',${JSON.stringify(item.name)})`,
          getIsSelected: item => state.filters.canal === item.name,
          clickTitle: 'Filtrar por canal',
          tooltipPrefix: 'Filtrar canal: '
        }
      );
    }
  }

  function renderTrend(rows){
    const title = document.getElementById('program-channel-trend-title');
    const host = document.getElementById('program-channel-trend');
    if(title) title.innerHTML = `Evolución por trimestre <span>${escapeHtml(state.filters.unidad)}</span>`;
    if(!host) return;
    const totals = new Map();
    rows.filter(row => periodSortValue(row.periodo) >= 0 && (state.mode !== 'Punto' || row.metrica !== 'balance')).forEach(row =>
      totals.set(row.periodo, (totals.get(row.periodo) || 0) + row.valor)
    );
    const points = [...totals.entries()]
      .map(([period,val]) => ({period,val}))
      .sort((a,b) => periodSortValue(a.period) - periodSortValue(b.period));
    if(!points.length) {
      host.innerHTML = '<div class="program-channel-empty">No hay periodos FY-Q para graficar.</div>';
      return;
    }

    const width = 680, height = 238, left = 68, right = 18, top = 18, bottom = 42;
    const graphWidth = width - left - right;
    const graphHeight = height - top - bottom;
    const maxValue = Math.max(...points.map(point => point.val), 1);
    const xFor = index => points.length === 1 ? left + graphWidth / 2 : left + index * graphWidth / (points.length - 1);
    const yFor = value => top + graphHeight - (value / maxValue) * graphHeight;
    const line = points.map((point,index) => `${xFor(index).toFixed(1)},${yFor(point.val).toFixed(1)}`).join(' ');
    const area = `${left},${top + graphHeight} ${line} ${xFor(points.length - 1)},${top + graphHeight}`;
    const labelStep = Math.max(1, Math.ceil(points.length / 8));
    let svg = `<svg viewBox="0 0 ${width} ${height}" class="program-channel-trend-svg" role="img" aria-label="Evolución por trimestre">`;
    svg += '<defs><linearGradient id="programTrendArea" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2ABFDF" stop-opacity=".28"/><stop offset="100%" stop-color="#2ABFDF" stop-opacity="0"/></linearGradient></defs>';
    [0,.25,.5,.75,1].forEach(step => {
      const y = top + graphHeight * (1 - step);
      svg += `<line x1="${left}" y1="${y}" x2="${width-right}" y2="${y}" stroke="var(--border)" stroke-width="1"/>`;
      svg += `<text x="${left-8}" y="${y+4}" text-anchor="end" fill="var(--text3)" font-size="9" font-family="Plus Jakarta Sans,sans-serif">${escapeHtml(formatCompact(maxValue * step, state.filters.unidad))}</text>`;
    });
    svg += `<polygon points="${area}" fill="url(#programTrendArea)"/>`;
    svg += `<polyline points="${line}" fill="none" stroke="#2ABFDF" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>`;
    points.forEach((point,index) => {
      const x = xFor(index), y = yFor(point.val);
      svg += `<circle cx="${x}" cy="${y}" r="4" fill="#06071A" stroke="#2ABFDF" stroke-width="2" data-tooltip="${escapeAttr(`${point.period}: ${formatMoney(point.val, state.filters.unidad)}`)}"/>`;
      if(index % labelStep === 0 || index === points.length - 1) {
        svg += `<text x="${x}" y="${height-13}" text-anchor="middle" fill="var(--text3)" font-size="9.5" font-family="Plus Jakarta Sans,sans-serif">${escapeHtml(point.period)}</text>`;
      }
    });
    svg += '</svg>';
    host.innerHTML = svg;
    if(typeof attachChartTooltips === 'function') attachChartTooltips(host);
  }

  function renderRanking(rows){
    const host = document.getElementById('program-channel-ranking');
    const title = document.getElementById('program-channel-ranking-title');
    if(!host) return;
    if(title) title.innerHTML = `Ranking de canales <span>${escapeHtml(state.filters.unidad)}</span>`;
    const executiveOverview = state.mode === 'Rebate' && state.filters.unidad === 'USD' &&
      !state.filters.canal && !state.filters.periodo && !state.filters.estado;
    let items = [];
    if(executiveOverview) {
      items = [
        { name:'Lenovo', val:BUSINESS_METRICS.lenovo.projection2026 },
        { name:'Dell', val:BUSINESS_METRICS.dell.approved },
        { name:'Epson', val:sumRows(rows.filter(row => row.canal === 'Epson')) }
      ];
    } else {
      const rankingRows = state.mode === 'Punto' ? rows.filter(row => row.metrica !== 'balance') : rows;
      const totals = new Map();
      rankingRows.forEach(row => totals.set(row.canal, (totals.get(row.canal) || 0) + row.valor));
      items = [...totals.entries()]
        .map(([name,val]) => ({ name,val }))
        .filter(item => item.val > 0)
        .sort((a,b) => b.val - a.val)
        .slice(0, 3);
    }
    if(!items.length) {
      host.innerHTML = '<div class="program-channel-empty">Sin valores confirmados para clasificar.</div>';
      return;
    }
    host.innerHTML = `<div class="program-channel-ranking">${items.map((item,index) => `
      <button type="button" class="program-channel-ranking-row" onclick="ProgramChannelModule.setFilter('canal',${JSON.stringify(item.name)})" title="Filtrar canal ${escapeAttr(item.name)}">
        <span class="program-channel-ranking-position">${index + 1}</span>
        <span class="program-channel-ranking-name">${escapeHtml(item.name)}</span>
        <span class="program-channel-ranking-value">${escapeHtml(state.mode === 'Punto' && state.filters.unidad === 'Puntos' ? formatInteger(item.val) : formatCompact(item.val, state.filters.unidad))}</span>
      </button>`).join('')}</div>`;
  }

  function insightCard(kind, icon, title, copy){
    return `<article class="program-channel-insight ${escapeAttr(kind)}">
      <span class="program-channel-insight-icon" aria-hidden="true">${escapeHtml(icon)}</span>
      <div><strong>${escapeHtml(title)}</strong><p>${escapeHtml(copy)}</p></div>
    </article>`;
  }

  function renderInsights(){
    const host = document.getElementById('program-channel-insights');
    if(!host) return;
    const channel = state.filters.canal;
    let cards = [];
    if(state.mode === 'Rebate') {
      const definitions = {
        Lenovo: insightCard('opportunity','★','Lenovo','2026 será el mejor año (~USD 82.535,60). WKS Expert es la oportunidad.'),
        Dell: insightCard('info','i','Dell','FY26 ejecutado y pagado; FY27 mantiene fondos aún no convertidos en ingreso.'),
        ASUS: insightCard('danger','!','ASUS','Solo 19% de cumplimiento de cuota: acción comercial prioritaria.'),
        Epson: insightCard('warning','•','Epson','Separar Consumo y Comercial; monitorear Pendiente vs. Procesado por mayorista.'),
        HPE: insightCard('info','i','HPE','La información validada disponible corresponde a puntos, no a rebates.'),
        Intel: insightCard('info','i','Intel','La información validada disponible corresponde a incentivos en puntos.')
      };
      cards = channel
        ? [definitions[channel]].filter(Boolean)
        : [definitions.Lenovo, definitions.Dell,
          insightCard('warning','!','Dell puntos','265 pendientes por redimir — acción requerida.'),
          definitions.ASUS];
    } else {
      const definitions = {
        Dell: insightCard('warning','!','Dell MyRewards','265 puntos pendientes por redimir — acción requerida por gerencia.'),
        HPE: insightCard('info','i','HPE Instant On','390 puntos disponibles; Visa Prepaid Card de USD 300 como equivalencia informativa.'),
        Lenovo: insightCard('opportunity','★','Incentivos SED','Plan Ultra Lenovo WS: 37 puntos y bono viajero COP separado.'),
        Intel: insightCard('warning','•','Intel','Programa activo sin valor confirmado; mantener en seguimiento.'),
        ASUS: insightCard('info','i','ASUS','No hay puntos validados; el seguimiento actual está en AGP Silver.'),
        Epson: insightCard('info','i','Epson','No hay puntos validados; los programas registrados corresponden a rebates.')
      };
      cards = channel
        ? [definitions[channel]].filter(Boolean)
        : [definitions.Dell, definitions.HPE, definitions.Lenovo,
          insightCard('warning','≠','Unidades separadas','Puntos y bono COP se consultan por unidad; nunca se suman.')];
    }
    host.innerHTML = cards.join('');
  }

  function renderTable(rows){
    const tableHost = document.getElementById('program-channel-table');
    const pagination = document.getElementById('program-channel-pagination');
    const meta = document.getElementById('program-channel-table-meta');
    const title = document.getElementById('program-channel-table-title');
    const exportButton = document.getElementById('program-channel-export-btn');
    if(!tableHost || !pagination) return;
    const pageCount = Math.max(1, Math.ceil(rows.length / state.pageSize));
    if(state.page > pageCount) state.page = pageCount;
    const start = (state.page - 1) * state.pageSize;
    const visible = rows.slice(start, start + state.pageSize);
    if(title) title.textContent = `Resumen ejecutivo por canal · ${state.mode === 'Punto' ? 'Puntos' : 'Rebates'}`;
    if(meta) meta.textContent = `${rows.length.toLocaleString('es-CO')} registros filtrados`;
    if(exportButton) exportButton.disabled = !rows.length;
    tableHost.innerHTML = `<table class="responsive-table program-channel-table">
      <thead><tr><th>Canal</th><th>Tipo</th><th>Programa</th><th>Periodo (FY-Q)</th><th>Valor</th><th>Unidad</th><th>Estado</th><th>Cliente / Ref</th></tr></thead>
      <tbody>${visible.length ? visible.map(row => {
        const group = statusGroup(row.estado);
        return `<tr>
          <td data-label="Canal" style="color:var(--text);font-weight:700">${escapeHtml(row.canal)}</td>
          <td data-label="Tipo"><span class="program-channel-type program-channel-type-${row.tipo.toLowerCase()}">${escapeHtml(row.tipo)}</span></td>
          <td data-label="Programa" class="program-channel-program" title="${escapeAttr(row.programa)}">${escapeHtml(row.programa)}</td>
          <td data-label="Periodo" class="td-mono">${escapeHtml(row.periodo)}</td>
          <td data-label="Valor" class="td-mono program-channel-value">${escapeHtml(row.unidad === 'Puntos' ? formatInteger(row.valor) : formatMoney(row.valor, row.unidad))}</td>
          <td data-label="Unidad">${escapeHtml(row.unidad)}</td>
          <td data-label="Estado"><span class="program-channel-state program-channel-state-${group}">${escapeHtml(row.estado)}</span></td>
          <td data-label="Cliente / Ref" class="program-channel-reference" title="${escapeAttr(row.clienteRef)}">${escapeHtml(row.clienteRef)}</td>
        </tr>`;
      }).join('') : '<tr><td colspan="8" class="program-channel-empty-cell">Sin registros para los filtros seleccionados.</td></tr>'}</tbody>
    </table>`;

    if(rows.length <= state.pageSize) {
      pagination.innerHTML = rows.length ? `<span>${start + 1}–${start + visible.length} de ${rows.length.toLocaleString('es-CO')}</span>` : '';
      return;
    }
    const pages = [];
    for(let page = Math.max(1, state.page - 2); page <= Math.min(pageCount, state.page + 2); page++) pages.push(page);
    pagination.innerHTML = `<span>${start + 1}–${start + visible.length} de ${rows.length.toLocaleString('es-CO')}</span>
      <div class="program-channel-page-actions">
        <button type="button" onclick="ProgramChannelModule.setPage(${state.page - 1})"${state.page === 1 ? ' disabled' : ''}>‹</button>
        ${pages.map(page => `<button type="button" class="${page === state.page ? 'active' : ''}" onclick="ProgramChannelModule.setPage(${page})">${page}</button>`).join('')}
        <button type="button" onclick="ProgramChannelModule.setPage(${state.page + 1})"${state.page === pageCount ? ' disabled' : ''}>›</button>
      </div>`;
  }

  function renderSourceStatus(){
    const host = document.getElementById('program-channel-source-status');
    if(!host) return;
    const sources = SOURCE_DEFINITIONS.map(def => state.sources[def.key]);
    const loaded = sources.filter(source => source.status === 'loaded');
    const missing = sources.filter(source => source.status === 'missing');
    const errors = sources.filter(source => source.status === 'error');
    const chips = `<span class="program-channel-source-chip program-channel-source-validated" title="Datos suministrados y validados por gerencia"><strong>Base validada</strong>${VALIDATED_RECORDS.length.toLocaleString('es-CO')} registros · ${escapeHtml(VALIDATED_AT)}</span>` + sources.map(source => {
      const label = source.fileName || source.label;
      const detail = source.status === 'loaded'
        ? `${source.records.length.toLocaleString('es-CO')} registros`
        : source.status === 'error' ? 'Error' : source.status === 'missing' ? 'No encontrado' : source.status === 'loading' ? 'Buscando' : 'Disponible al actualizar';
      return `<span class="program-channel-source-chip program-channel-source-${source.status}" title="${escapeAttr(source.error || label)}"><strong>${escapeHtml(source.label)}</strong>${escapeHtml(detail)}</span>`;
    }).join('');
    let summary = state.loading ? 'Base validada visible · buscando y normalizando reportes...' : `Base gerencial activa · ${loaded.length} de ${sources.length} fuentes en vivo`;
    if(missing.length) summary += ` · Faltan: ${missing.map(source => source.label).join(', ')}`;
    if(errors.length) summary += ` · ${errors.length} fuente(s) con error`;
    if(state.error) summary = state.error;
    host.innerHTML = `<div class="program-channel-source-summary">${escapeHtml(summary)}</div><div class="program-channel-source-chips">${chips}</div>`;
  }

  function render(){
    const page = document.getElementById('page-programas');
    if(!page || !canAccess()) return;
    const pointButton = document.getElementById('program-channel-mode-points');
    const rebateButton = document.getElementById('program-channel-mode-rebates');
    if(pointButton) {
      pointButton.classList.toggle('active', state.mode === 'Punto');
      pointButton.setAttribute('aria-selected', state.mode === 'Punto' ? 'true' : 'false');
    }
    if(rebateButton) {
      rebateButton.classList.toggle('active', state.mode === 'Rebate');
      rebateButton.setAttribute('aria-selected', state.mode === 'Rebate' ? 'true' : 'false');
    }
    syncFilters();
    renderSourceStatus();
    const note = document.getElementById('program-channel-unit-note');
    if(note) {
      note.textContent = state.mode === 'Rebate'
        ? `Vista activa: rebates en ${state.filters.unidad}. USD y COP nunca se suman entre sí.`
        : state.filters.unidad === 'COP'
          ? 'Vista activa: incentivos monetarios en COP. No se suman con los puntos.'
          : 'Vista activa: puntos. No se suman con rebates ni incentivos monetarios.';
    }
    const rows = getVisibleRows();
    renderKpis();
    renderBarsChart(rows);
    renderTrend(rows);
    renderRanking(rows);
    renderInsights();
    renderTable(rows);
  }

  async function exportExcel(){
    const rows = getVisibleRows();
    if(!rows.length) return;
    const button = document.getElementById('program-channel-export-btn');
    const original = button ? button.innerHTML : '';
    if(button) {
      button.disabled = true;
      button.innerHTML = '<span class="export-excel-icon">…</span><span>Generando</span>';
    }
    try {
      if(typeof ensureExcelJsForExport === 'function') await ensureExcelJsForExport();
      if(typeof ExcelJS === 'undefined') throw new Error('ExcelJS no esta disponible.');
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Forecast 2026 - Provexpress';
      workbook.created = new Date();
      const worksheet = workbook.addWorksheet('Tabla Maestra');
      worksheet.views = [{ state:'frozen', ySplit:1 }];
      worksheet.columns = [
        { header:'Canal', key:'canal', width:14 },
        { header:'Tipo(Rebate/Punto)', key:'tipo', width:22 },
        { header:'Programa', key:'programa', width:52 },
        { header:'Periodo(FY-Q)', key:'periodo', width:16 },
        { header:'Valor', key:'valor', width:18 },
        { header:'Unidad(USD/COP/Puntos)', key:'unidad', width:24 },
        { header:'Estado', key:'estado', width:24 },
        { header:'Cliente/Ref', key:'clienteRef', width:36 }
      ];
      rows.forEach(row => worksheet.addRow({
        canal: row.canal,
        tipo: row.tipo,
        programa: row.programa,
        periodo: row.periodo,
        valor: row.valor,
        unidad: row.unidad,
        estado: row.estado,
        clienteRef: row.clienteRef
      }));
      worksheet.autoFilter = { from:'A1', to:'H1' };
      worksheet.getRow(1).height = 26;
      worksheet.getRow(1).font = { name:'Aptos Display', size:10, bold:true, color:{ argb:'FFFFFFFF' } };
      worksheet.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0B1956' } };
      worksheet.getRow(1).alignment = { vertical:'middle', horizontal:'center' };
      for(let index = 2; index <= rows.length + 1; index++) {
        const row = worksheet.getRow(index);
        row.font = { name:'Aptos', size:9, color:{ argb:'FF172033' } };
        row.alignment = { vertical:'middle' };
        row.getCell(3).alignment = { vertical:'middle', wrapText:true };
        row.getCell(5).numFmt = rows[index - 2].unidad === 'Puntos' ? '#,##0' : '#,##0.00';
        row.getCell(5).alignment = { horizontal:'right' };
        if(index % 2 === 0) row.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF5F7FD' } };
      }
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      const stamp = typeof getBogotaTimestampForFile === 'function'
        ? getBogotaTimestampForFile()
        : new Date().toISOString().replace(/[:.]/g,'-');
      link.download = `Rebates_Puntos_Incentivos_${state.mode}_${stamp}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    } catch(error) {
      console.error('[PROGRAMAS EXPORT]', error);
      alert('No se pudo exportar el detalle: ' + (error.message || error));
    } finally {
      if(button) {
        button.innerHTML = original;
        button.disabled = false;
      }
    }
  }

  window.ProgramChannelModule = {
    canAccess,
    loadFromSharePoint,
    reloadFromSharePoint,
    openLocalFiles,
    handleLocalFiles,
    setMode,
    setFilter,
    clearFilters,
    setPage,
    exportExcel,
    render,
    parseWorkbook,
    getData: () => state.data.slice(),
    getState: () => ({
      mode: state.mode,
      filters: { ...state.filters },
      loading: state.loading,
      error: state.error,
      sourceSummary: Object.fromEntries(SOURCE_DEFINITIONS.map(def => [def.key, {
        status: state.sources[def.key].status,
        fileName: state.sources[def.key].fileName,
        records: state.sources[def.key].records.length
      }]))
    })
  };
})();
