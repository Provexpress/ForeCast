/* ══════════════════════════════════════
   REBATES, PUNTOS E INCENTIVOS POR CANAL
   Normalizacion de reportes de fabricantes
══════════════════════════════════════ */
(function(){
  'use strict';

  const SHAREPOINT_FOLDER_NAME = 'Fabricantes';
  const EMBEDDED_REPORT_FILE = 'Informe_Puntos_Incentivos 1.xlsx';
  const EMBEDDED_REPORT_URL = 'Informe_Puntos_Incentivos%201.xlsx';

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

  const CHANNELS = ['Dell','Lenovo','HPE','ASUS','Epson','Intel','Microsoft'];
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
    workbook: {
      status: 'idle',
      fileName: EMBEDDED_REPORT_FILE,
      error: '',
      sheets: [],
      activeSheet: '',
      search: ''
    },
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
    const match = compact.match(/FY(\d{4}|\d{2})-?Q([1-4])/);
    if(match) {
      const year = match[1].length === 4 ? match[1].slice(-2) : match[1];
      return `FY${year}-Q${match[2]}`;
    }
    const reverse = compact.match(/Q([1-4])-?FY(\d{4}|\d{2})/);
    if(reverse) {
      const year = reverse[2].length === 4 ? reverse[2].slice(-2) : reverse[2];
      return `FY${year}-Q${reverse[1]}`;
    }
    return text === 'SIN PERIODO' ? 'Sin periodo' : text;
  }

  function periodFromParts(yearValue, quarterValue){
    const yearMatch = normalizeText(yearValue).toUpperCase().match(/(?:FY)?(\d{4}|\d{2})/);
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
    const yearMatch = text.match(/FY(\d{4}|\d{2})/);
    const quarterMatch = normalizeText(quarter).toUpperCase().match(/Q([1-4])/) || text.match(/([1-4])Q/);
    if(!yearMatch || !quarterMatch) return 'Sin periodo';
    const rawYear = yearMatch[1];
    const year = rawYear.length === 4 ? rawYear.slice(-2) : rawYear;
    return `FY${year}-Q${quarterMatch[1]}`;
  }

  function periodSortValue(value){
    const match = normalizeText(value).match(/^FY(\d{2})-Q([1-4])$/i);
    if(match) return Number(match[1]) * 100 + Number(match[2]) * 3;
    const monthMatch = normalizeText(value).match(/^FY(\d{2})-M(0?[1-9]|1[0-2])$/i);
    if(monthMatch) return Number(monthMatch[1]) * 100 + Number(monthMatch[2]);
    return -1;
  }

  function periodFromMonth(yearValue, monthValue){
    const year = Number(yearValue);
    const month = Number(monthValue);
    if(!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return 'Sin periodo';
    return `FY${String(year).slice(-2)}-M${String(month).padStart(2,'0')}`;
  }

  function programFiscalPeriod(program, quarter, fallbackDate){
    const programText = normalizeText(program).toUpperCase();
    const yearMatch = programText.match(/FY(\d{4}|\d{2})/);
    const quarterMatch = normalizeText(quarter).toUpperCase().match(/Q([1-4])/);
    if(yearMatch && quarterMatch) {
      const year = yearMatch[1].length === 4 ? yearMatch[1].slice(-2) : yearMatch[1];
      return `FY${year}-Q${quarterMatch[1]}`;
    }
    return periodFromDate(fallbackDate);
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
    const sourceDate = parseDate(data.fecha);
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
      metrica: data.metrica === 'balance' ? 'balance' : 'movement',
      movimiento: normalizeText(data.movimiento),
      fecha: sourceDate ? `${sourceDate.getFullYear()}-${String(sourceDate.getMonth() + 1).padStart(2,'0')}-${String(sourceDate.getDate()).padStart(2,'0')}` : '',
      anio: Number(data.anio) || (sourceDate ? sourceDate.getFullYear() : 0),
      mes: Number(data.mes) || (sourceDate ? sourceDate.getMonth() + 1 : 0),
      esEquivalencia: Boolean(data.esEquivalencia)
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
          movimiento: 'rebate',
          programa: program,
          periodo: period,
          valor: row[4],
          unidad: 'USD',
          estado: stateValue,
          clienteRef: reference,
          fecha: row[1],
          hoja: epsonName
        }));
        const copValue = toNumber(row[7]) || toNumber(row[10]);
        if(copValue > 0) {
          records.push(createRecord('platforms', {
            canal: 'Epson',
            tipo: 'Rebate',
            movimiento: 'rebate',
            programa: program,
            periodo: period,
            valor: copValue,
            unidad: 'COP',
            estado: stateValue,
            clienteRef: reference,
            fecha: row[1],
            esEquivalencia: !toNumber(row[7]),
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
      const conversionRate = toNumber((hpeInstantRows[hpeHeader] || [])[6]);
      let currentRedemptionsUsd = 0;
      let currentRedemptionsPeriod = 'Sin periodo';
      let redemptionCount = 0;
      for(let rowIndex = hpeHeader + 1; rowIndex < hpeInstantRows.length; rowIndex++) {
        const row = hpeInstantRows[rowIndex] || [];
        const program = valueByHeader(row, indexes, ['Descripción','Descripcion']);
        const value = toNumber(valueByHeader(row, indexes, ['Incentivos redimidos']));
        if(!normalizeText(program) || value <= 0) continue;
        const extractedPeriod = normalizeText(program).match(/FY\d{2}Q[1-4]/i);
        const period = extractedPeriod ? extractedPeriod[0] : periodFromDate(valueByHeader(row, indexes, ['Fecha']));
        const receipt = normalizeText(valueByHeader(row, indexes, ['Recibo de pago']));
        const redemptionId = normalizeText(valueByHeader(row, indexes, ['Identificación de canje','Identificacion de canje']));
        if(redemptionCount < 4) {
          currentRedemptionsUsd += value;
          if(redemptionCount === 0) currentRedemptionsPeriod = period;
        }
        redemptionCount++;
        records.push(createRecord('platforms', {
          canal: 'HPE',
          tipo: 'Punto',
          movimiento: 'canje',
          programa: program,
          periodo: period,
          valor: value,
          unidad: 'USD',
          estado: receipt ? 'Redimido / Soporte registrado' : 'Redimido / Soporte pendiente',
          clienteRef: [redemptionId, receipt || 'Recibo pendiente'].filter(Boolean).join(' · '),
          fecha: valueByHeader(row, indexes, ['Fecha']),
          hoja: hpeInstantName
        }));
      }
      if(currentRedemptionsUsd > 0 && conversionRate > 0) {
        records.push(createRecord('platforms', {
          canal:'HPE', tipo:'Punto', movimiento:'canje_resumen', metrica:'balance',
          programa:'HPE Instant On · Redenciones vigentes', periodo:currentRedemptionsPeriod,
          valor:currentRedemptionsUsd * conversionRate, unidad:'COP', estado:'Redimido',
          clienteRef:`${formatInteger(currentRedemptionsUsd)} USD × TRM ${formatInteger(conversionRate)}`,
          esEquivalencia:true, hoja:hpeInstantName
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
        const cardUsd = toNumber(row[1]);
        const cardCop = toNumber(row[3]);
        const available = toNumber(row[4]);
        const deliveredDate = parseDate(row[2]);
        const cardState = deliveredDate && deliveredDate.getTime() <= Date.now() ? 'Entregado' : 'Pendiente de entrega';
        if(cardUsd > 0) {
          records.push(createRecord('platforms', {
            canal:'HPE', tipo:'Punto', movimiento:'tarjeta', programa:concept, periodo:period,
            valor:cardUsd, unidad:'USD', estado:cardState, fecha:row[2],
            clienteRef:'Physical Visa Prepaid Card', hoja:hpeName
          }));
        }
        if(cardCop > 0) records.push(createRecord('platforms', {
          canal:'HPE', tipo:'Punto', movimiento:'tarjeta', programa:concept, periodo:period,
          valor:cardCop, unidad:'COP', estado:cardState, fecha:row[2],
          clienteRef:`Equivalencia de ${formatMoney(cardUsd, 'USD')}`, esEquivalencia:true, hoja:hpeName
        }));
        if(available > 0) {
          records.push(createRecord('platforms', {
            canal: 'HPE', tipo: 'Punto', programa: 'Saldo disponible HPE', periodo: period,
            movimiento:'saldo_disponible', metrica: 'balance', valor: available, unidad: 'Puntos', estado: 'Disponible',
            clienteRef: concept, fecha:row[2], hoja: hpeName
          }));
        }
      }
    }

    const lenovoCardName = findSheetName(workbook, key => key === 'lenovo');
    const lenovoCardRows = sheetRows(workbook, lenovoCardName);
    const lenovoCardHeader = findHeaderIndex(lenovoCardRows, ['Fecha','Descripcion','Debito USD','Credito USD']);
    if(lenovoCardHeader >= 0) {
      const indexes = headerIndexes(lenovoCardRows[lenovoCardHeader]);
      for(let rowIndex = lenovoCardHeader + 1; rowIndex < lenovoCardRows.length; rowIndex++) {
        const row = lenovoCardRows[rowIndex] || [];
        const dateValue = valueByHeader(row, indexes, ['Fecha']);
        if(!parseDate(dateValue)) continue;
        const description = normalizeText(valueByHeader(row, indexes, ['Descripcion','Descripción']));
        const debit = toNumber(valueByHeader(row, indexes, ['Debito USD','Débito USD']));
        const credit = toNumber(valueByHeader(row, indexes, ['Credito USD','Crédito USD']));
        const copValue = toNumber(valueByHeader(row, indexes, ['Cop']));
        const period = periodFromDate(dateValue);
        let movement = 'otro';
        let label = 'Movimiento';
        let status = 'Registrado';
        if(credit > 0) { movement = 'recarga'; label = 'Recarga'; status = 'Recargado'; }
        else if(/^fee general credit/i.test(description)) { movement = 'comision_recarga'; label = 'Comisión recarga'; status = 'Comisión'; }
        else if(/^fee purchase/i.test(description)) { movement = 'comision_pos'; label = 'Comisión POS'; status = 'Comisión'; }
        else if(debit > 0) { movement = 'compra'; label = 'Compra'; status = 'Consumido'; }
        const usdValue = credit > 0 ? credit : debit;
        if(usdValue > 0) records.push(createRecord('platforms', {
          grupo:'lenovo_card', canal:'Lenovo', tipo:'Punto', movimiento:movement,
          programa:`Lenovo tarjeta · ${label}`, periodo:period, valor:usdValue, unidad:'USD', estado:status,
          clienteRef:description || label, fecha:dateValue, hoja:lenovoCardName
        }));
        if(copValue > 0 && debit > 0) records.push(createRecord('platforms', {
          grupo:'lenovo_card', canal:'Lenovo', tipo:'Punto', movimiento:movement,
          programa:`Lenovo tarjeta · ${label}`, periodo:period, valor:copValue, unidad:'COP', estado:status,
          clienteRef:description || label, fecha:dateValue, esEquivalencia:true, hoja:lenovoCardName
        }));
      }

      const metricRow = label => lenovoCardRows.find(row => normalizeKey((row || [])[9]) === normalizeKey(label)) || [];
      [
        { label:'Total recargas', movement:'resumen_recargas' },
        { label:'Compras identificadas', movement:'resumen_compras' },
        { label:'Comisiones POS', movement:'resumen_comision_pos' },
        { label:'Comisiones recarga', movement:'resumen_comision_recarga' },
        { label:'Saldo teórico', movement:'saldo_teorico' },
        { label:'Diferencia', movement:'diferencia' }
      ].forEach(metric => {
        const row = metricRow(metric.label);
        const value = toNumber(row[10]);
        if(value <= 0) return;
        records.push(createRecord('platforms', {
          grupo:'lenovo_card', canal:'Lenovo', tipo:'Punto', movimiento:metric.movement, metrica:'balance',
          programa:`Lenovo tarjeta · ${metric.label}`, periodo:'Sin periodo', valor:value, unidad:'USD', estado:'Resumen',
          clienteRef:'Total explícito de la hoja Lenovo', hoja:lenovoCardName
        }));
      });
      const availableRow = metricRow('Saldo disponible');
      const availableUsd = toNumber(availableRow[10]);
      const availableCop = toNumber(availableRow[11]);
      if(availableUsd > 0) records.push(createRecord('platforms', {
        grupo:'lenovo_card', canal:'Lenovo', tipo:'Punto', movimiento:'saldo_disponible', metrica:'balance',
        programa:'Lenovo tarjeta · Saldo disponible', periodo:'Sin periodo', valor:availableUsd, unidad:'USD', estado:'Disponible',
        clienteRef:'Saldo actual de tarjeta', hoja:lenovoCardName
      }));
      if(availableCop > 0) records.push(createRecord('platforms', {
        grupo:'lenovo_card', canal:'Lenovo', tipo:'Punto', movimiento:'saldo_disponible', metrica:'balance',
        programa:'Lenovo tarjeta · Saldo disponible', periodo:'Sin periodo', valor:availableCop, unidad:'COP', estado:'Disponible',
        clienteRef:`Equivalencia de ${formatMoney(availableUsd, 'USD')}`, esEquivalencia:true, hoja:lenovoCardName
      }));

      [2025, 2026].forEach(year => {
        const row = lenovoCardRows.find(item => normalizeKey((item || [])[4]) === String(year)) || [];
        const value = toNumber(row[6]);
        if(value <= 0) return;
        records.push(createRecord('platforms', {
          grupo:'lenovo_card', canal:'Lenovo', tipo:'Punto', movimiento:`consumo_${year}`, metrica:'balance',
          programa:`Lenovo tarjeta · Consumo ${year}`, periodo:`FY${String(year).slice(-2)}`, valor:value, unidad:'COP', estado:'Resumen',
          clienteRef:'Total anual explícito de la hoja Lenovo', anio:year, hoja:lenovoCardName
        }));
      });

      [
        { label:'Saldo global', movement:'saldo_global', state:'Emitido' },
        { label:'Pendientes por asignar', movement:'pendiente_asignar', state:'Pendiente' },
        { label:'saldo', movement:'saldo_asignado', state:'Disponible' }
      ].forEach(metric => {
        const row = metricRow(metric.label);
        const value = toNumber(row[10]);
        if(value <= 0) return;
        records.push(createRecord('platforms', {
          grupo:'lenovo_card', canal:'Lenovo', tipo:'Punto', movimiento:metric.movement, metrica:'balance',
          programa:`Lenovo puntos · ${metric.label}`, periodo:'Sin periodo', valor:value, unidad:'Puntos', estado:metric.state,
          clienteRef:'Control de asignación Lenovo', hoja:lenovoCardName
        }));
      });
    }

    const lenovoProgramsName = findSheetName(workbook, key => key === 'lenovo programa de canales');
    const lenovoProgramsRows = sheetRows(workbook, lenovoProgramsName);
    let currentLenovoQuarter = '';
    let lenovoPaymentBlock = -1;
    const lenovoPaymentYears = [2023, 2024, 2025, 2026];
    for(let rowIndex = 0; rowIndex < lenovoProgramsRows.length; rowIndex++) {
      const row = lenovoProgramsRows[rowIndex] || [];
      const quarter = normalizeText(row[0]);
      if(normalizeKey(quarter) === 'program start date') {
        lenovoPaymentBlock++;
        currentLenovoQuarter = '';
        continue;
      }
      if(/^Q[1-4]_/.test(quarter)) currentLenovoQuarter = quarter;
      const program = normalizeText(row[1]);
      const value = toNumber(row[2]);
      const paymentYear = lenovoPaymentYears[lenovoPaymentBlock];
      if(!paymentYear || !program || !/[A-Za-z]/.test(program) || value <= 0 || /total general/i.test(program)) continue;
      records.push(createRecord('platforms', {
        grupo:'lenovo', canal:'Lenovo', tipo:'Rebate', movimiento:'pago',
        programa:program, periodo:periodFromParts(paymentYear, currentLenovoQuarter), valor:value, unidad:'USD', estado:'Pagado',
        clienteRef:`Tabla dinámica · ${currentLenovoQuarter.replace(/_/g,' ')} · pago confirmado`, hoja:lenovoProgramsName
      }));
    }

    const microsoftName = findSheetName(workbook, key => key === 'microsoft');
    const microsoftRows = sheetRows(workbook, microsoftName);
    const microsoftHeader = findHeaderIndex(microsoftRows, ['participantID','programName','earned','paymentStatus']);
    if(microsoftHeader >= 0) {
      const indexes = headerIndexes(microsoftRows[microsoftHeader]);
      for(let rowIndex = microsoftHeader + 1; rowIndex < microsoftRows.length; rowIndex++) {
        const row = microsoftRows[rowIndex] || [];
        const participantId = valueByHeader(row, indexes, ['participantID']);
        const program = normalizeText(valueByHeader(row, indexes, ['programName']));
        const status = normalizeText(valueByHeader(row, indexes, ['paymentStatus']));
        if(!participantId || !program || !status) continue;
        const earnedCop = toNumber(valueByHeader(row, indexes, ['earned']));
        const earnedUsd = toNumber(valueByHeader(row, indexes, ['earnedUSD']));
        const paidCop = toNumber(valueByHeader(row, indexes, ['totalPayment']));
        const dateValue = valueByHeader(row, indexes, ['paymentDat','paymentDate']);
        const parsedDate = parseDate(dateValue);
        const year = toNumber(valueByHeader(row, indexes, ['año','ano'])) || (parsedDate ? parsedDate.getFullYear() : 0);
        const month = toNumber(valueByHeader(row, indexes, ['mes'])) || (parsedDate ? parsedDate.getMonth() + 1 : 0);
        const period = periodFromMonth(year, month);
        const statusKey = normalizeKey(status);
        const movement = statusKey.includes('sent') ? 'pago' : statusKey.includes('upcoming') ? 'proximo' : 'en_riesgo';
        const paymentId = normalizeText(valueByHeader(row, indexes, ['paymentID']));
        const description = normalizeText(valueByHeader(row, indexes, ['paymentStatusDescription']));
        const reference = [`MPN ${participantId}`, paymentId ? `Pago ${paymentId}` : '', description].filter(Boolean).join(' · ');
        records.push(createRecord('platforms', {
          grupo:'microsoft', canal:'Microsoft', tipo:'Rebate', movimiento:movement, programa:program, periodo:period,
          valor:paidCop || earnedCop, unidad:'COP', estado:status, clienteRef:reference,
          fecha:dateValue, anio:year, mes:month, hoja:microsoftName
        }));
        if(earnedUsd > 0) records.push(createRecord('platforms', {
          grupo:'microsoft', canal:'Microsoft', tipo:'Rebate', movimiento:movement, programa:program, periodo:period,
          valor:earnedUsd, unidad:'USD', estado:status, clienteRef:reference,
          fecha:dateValue, anio:year, mes:month, esEquivalencia:true, hoja:microsoftName
        }));
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
        movimiento: 'pendiente',
        programa: `Incentivo Intel de ${formatInteger(denomination)} puntos`,
        periodo: 'Sin periodo',
        valor: value,
        unidad: 'Puntos',
        estado: intelState,
        clienteRef: `Cantidad: ${formatInteger(quantity)}`,
        hoja: intelName
      }));
      const copValue = toNumber(row[3]);
      if(copValue > 0) records.push(createRecord('platforms', {
        canal:'Intel', tipo:'Punto', metrica:'balance', movimiento:'pendiente',
        programa:`Incentivo Intel de ${formatInteger(denomination)} puntos`, periodo:'Sin periodo',
        valor:copValue, unidad:'COP', estado:intelState,
        clienteRef:`Equivalencia · cantidad ${formatInteger(quantity)}`, esEquivalencia:true, hoja:intelName
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

  function isEmbeddedReportFile(fileName){
    const key = normalizeFileName(fileName);
    return key.includes('informe puntos incentivos') || key.includes('puntos incentivos 1');
  }

  function isEmptyWorkbookCell(cell){
    return cell == null || normalizeText(cell.text) === '';
  }

  function formatWorkbookCell(cell){
    if(!cell) return '';
    const value = cell.v;
    if(value == null) return '';
    if(cell.t === 'd' && value instanceof Date) {
      return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2,'0')}-${String(value.getDate()).padStart(2,'0')}`;
    }
    if(cell.t === 'n' && cell.z && /[ymd]/i.test(cell.z) && typeof XLSX !== 'undefined' && XLSX.SSF && XLSX.SSF.parse_date_code) {
      const parsedDate = XLSX.SSF.parse_date_code(value);
      if(parsedDate && parsedDate.y && parsedDate.m && parsedDate.d) {
        return `${parsedDate.y}-${String(parsedDate.m).padStart(2,'0')}-${String(parsedDate.d).padStart(2,'0')}`;
      }
    }
    if(typeof value === 'number') {
      if(Number.isInteger(value)) return value.toLocaleString('es-CO', { maximumFractionDigits:0 });
      return value.toLocaleString('es-CO', { minimumFractionDigits:2, maximumFractionDigits:12 });
    }
    return String(value);
  }

  function workbookSheetToRows(workbook, sheetName){
    const sheet = workbook.Sheets && workbook.Sheets[sheetName];
    if(!sheet || !sheet['!ref']) return { rows: [], merges: [] };
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const rows = [];
    for(let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
      const row = [];
      for(let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        const rawCell = sheet[address];
        const text = formatWorkbookCell(rawCell);
        row.push({
          text,
          raw: rawCell && rawCell.v != null ? rawCell.v : null,
          type: rawCell && rawCell.t || '',
          formula: rawCell && rawCell.f || ''
        });
      }
      rows.push(row);
    }
    while(rows.length && rows[rows.length - 1].every(isEmptyWorkbookCell)) rows.pop();
    let lastCol = 0;
    rows.forEach(row => {
      for(let colIndex = row.length - 1; colIndex >= 0; colIndex--) {
        if(!isEmptyWorkbookCell(row[colIndex])) {
          lastCol = Math.max(lastCol, colIndex + 1);
          break;
        }
      }
    });
    const clippedRows = rows.map(row => row.slice(0, lastCol));
    const merges = (sheet['!merges'] || []).map(merge => ({
      startRow: merge.s.r - range.s.r,
      startCol: merge.s.c - range.s.c,
      endRow: merge.e.r - range.s.r,
      endCol: merge.e.c - range.s.c
    })).filter(merge =>
      merge.startRow >= 0 &&
      merge.startCol >= 0 &&
      merge.startRow < clippedRows.length &&
      merge.startCol < lastCol
    ).map(merge => ({
      startRow: merge.startRow,
      startCol: merge.startCol,
      endRow: Math.min(merge.endRow, clippedRows.length - 1),
      endCol: Math.min(merge.endCol, lastCol - 1)
    }));
    return { rows: clippedRows, merges };
  }

  function parseEmbeddedReportWorkbook(workbook, fileName){
    const sheets = (workbook.SheetNames || []).map(name => ({
      name,
      ...workbookSheetToRows(workbook, name)
    }));
    state.workbook = {
      status: 'loaded',
      fileName: fileName || EMBEDDED_REPORT_FILE,
      error: '',
      sheets,
      activeSheet: state.workbook.activeSheet && sheets.some(sheet => sheet.name === state.workbook.activeSheet)
        ? state.workbook.activeSheet
        : (sheets[0] && sheets[0].name || ''),
      search: state.workbook.search || ''
    };
  }

  async function loadEmbeddedReportFromBuffer(buffer, fileName){
    if(typeof XLSX === 'undefined') throw new Error('La libreria XLSX no esta disponible.');
    const workbook = XLSX.read(buffer, { type:'array', cellDates:true, cellNF:true, cellText:false });
    parseEmbeddedReportWorkbook(workbook, fileName);
  }

  async function reloadEmbeddedReport(){
    if(!canAccess() || state.workbook.status === 'loading') return;
    state.workbook.status = 'loading';
    state.workbook.error = '';
    renderWorkbookReport();
    try {
      const response = await fetch(EMBEDDED_REPORT_URL, { cache:'no-store' });
      if(!response.ok) throw new Error(`No se pudo leer ${EMBEDDED_REPORT_FILE} (${response.status}).`);
      await loadEmbeddedReportFromBuffer(await response.arrayBuffer(), EMBEDDED_REPORT_FILE);
    } catch(error) {
      state.workbook.status = 'error';
      state.workbook.error = `${error.message || error}. Usa Cargar Excel y selecciona ${EMBEDDED_REPORT_FILE}.`;
      console.warn('[PROGRAMAS REPORT]', error);
    }
    renderWorkbookReport();
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
    const live = SOURCE_DEFINITIONS.flatMap(def => state.sources[def.key].records || []);
    const loadedGroups = new Set(live.map(row => row.grupo).filter(Boolean));
    const validated = VALIDATED_RECORDS.filter(row => !loadedGroups.has(row.grupo));
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
      if(isEmbeddedReportFile(file.name)) {
        try {
          state.workbook.status = 'loading';
          renderWorkbookReport();
          await loadEmbeddedReportFromBuffer(await file.arrayBuffer(), file.name);
          recognized.push('embedded-report');
        } catch(error) {
          state.workbook.status = 'error';
          state.workbook.fileName = file.name;
          state.workbook.error = error.message || String(error);
        }
        continue;
      }
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

  function setWorkbookSheet(sheetName){
    if(!state.workbook.sheets.some(sheet => sheet.name === sheetName)) return;
    state.workbook.activeSheet = sheetName;
    renderWorkbookReport();
  }

  function setWorkbookSearch(value){
    state.workbook.search = normalizeText(value);
    renderWorkbookReport();
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
    const rows = rowsWithoutUnit().filter(row => !state.filters.unidad || row.unidad === state.filters.unidad);
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
    const allowedUnits = state.mode === 'Punto' ? ['Puntos','USD','COP'] : ['USD','COP'];
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
    if(typeLabel) typeLabel.textContent = state.mode === 'Punto' ? 'Punto / Incentivo' : 'Rebate';
  }

  function statusGroup(value){
    const key = normalizeKey(value);
    if(/expir/.test(key)) return 'expired';
    if(/declin|cancel|rechaz|perdid|forfeit/.test(key)) return 'rejected';
    if(/pend|upcoming|review|released|dispon|unclaimed|generad|vigent|oportunidad|alerta/.test(key)) return 'pending';
    if(/redimid|asignad|allocated|cargado|entregad|recargad|consumid/.test(key)) return 'redeemed';
    if(/pagad|paid|procesad|complete|sent/.test(key)) return 'paid';
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
    if(unit === 'Puntos') return formatInteger(value);
    return formatMoney(value, unit);
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

    if(state.mode === 'Punto') {
      const unit = state.filters.unidad;
      const unitRows = allUnitsRows.filter(row => row.unidad === unit);

      if(!hasDetailFilters && state.filters.canal === 'Dell' && unit === 'Puntos') {
        const movements = unitRows.filter(row => row.metrica !== 'balance');
        host.innerHTML = [
          kpiCard('Puntos históricos', formatInteger(sumRows(movements)), 'FY22-Q1 a FY23-Q2', '#2ABFDF'),
          kpiCard('Pendientes por redimir', formatInteger(BUSINESS_METRICS.dell.unredeemedPoints), 'Acción requerida', '#F0A020', 'program-channel-kpi-alert'),
          kpiCard('Estado clave', 'Unclaimed', 'Dell MyRewards', '#E84040'),
          kpiCard('Unidad', 'Puntos', 'Nunca se suman con dinero', '#8B5FC8')
        ].join('');
        return;
      }

      if(!hasDetailFilters && state.filters.canal === 'HPE' && unit === 'Puntos') {
        const available = sumRows(unitRows.filter(row => row.movimiento === 'saldo_disponible')) || BUSINESS_METRICS.hpe.availablePoints;
        const sedPoints = sumRows(unitRows.filter(row => row.grupo === 'sed' && row.valor > 0));
        host.innerHTML = [
          kpiCard('Saldo disponible HPE', formatInteger(available), 'Puntos pendientes de canje', '#2ABFDF'),
          kpiCard('Incentivos SED', formatInteger(sedPoints), 'Puntos vigentes separados', '#8B5FC8'),
          kpiCard('Tarjeta física', formatMoney(BUSINESS_METRICS.hpe.visaUsd, 'USD'), 'Entregada · se consulta en USD', '#0DBF82'),
          kpiCard('Control de unidad', 'Puntos', 'Sin sumar USD ni COP', '#F0A020')
        ].join('');
        return;
      }

      if(!hasDetailFilters && state.filters.canal === 'HPE' && unit === 'USD') {
        const redemptions = unitRows.filter(row => row.movimiento === 'canje');
        const card = unitRows.filter(row => row.movimiento === 'tarjeta');
        const currentRedemptions = sumRows(redemptions.slice(0, 4));
        const missingSupport = redemptions.filter(row => /soporte pendiente/i.test(row.estado)).length;
        host.innerHTML = [
          kpiCard('Redenciones acumuladas', formatMoney(sumRows(redemptions), 'USD'), `${redemptions.length} canjes HPE Instant On`, '#2ABFDF'),
          kpiCard('Canjes vigentes', formatMoney(currentRedemptions, 'USD'), 'Últimos 4 registros del control', '#8B5FC8'),
          kpiCard('Tarjeta Visa física', formatMoney(sumRows(card), 'USD'), 'Entregada el 1 jul 2026', '#0DBF82'),
          kpiCard('Soportes pendientes', formatInteger(missingSupport), 'Recibos por adjuntar', '#E84040', missingSupport ? 'program-channel-kpi-alert' : '')
        ].join('');
        return;
      }

      if(!hasDetailFilters && state.filters.canal === 'HPE' && unit === 'COP') {
        const cardCop = sumRows(unitRows.filter(row => row.movimiento === 'tarjeta'));
        const redemptionCop = sumRows(unitRows.filter(row => row.movimiento === 'canje_resumen'));
        host.innerHTML = [
          kpiCard('Tarjeta Visa física', formatMoney(cardCop, 'COP'), 'USD 300 × TRM 3.300 del Excel', '#0DBF82'),
          kpiCard('Canjes vigentes', formatMoney(redemptionCop, 'COP'), 'USD 828 × TRM 3.600 del Excel', '#2ABFDF'),
          kpiCard('Incentivos visibles', formatMoney(cardCop + redemptionCop, 'COP'), 'Equivalencias, sin duplicar USD', '#8B5FC8'),
          kpiCard('Dato fuente tarjeta', '$ 990.000', 'Valor calculado en la hoja HPE', '#F0A020')
        ].join('');
        return;
      }

      if(!hasDetailFilters && state.filters.canal === 'Lenovo' && unit === 'Puntos') {
        const metric = movement => sumRows(unitRows.filter(row => row.movimiento === movement));
        const sedPoints = sumRows(unitRows.filter(row => row.grupo === 'sed' && row.valor > 0));
        host.innerHTML = [
          kpiCard('Saldo global / emisiones', formatInteger(metric('saldo_global')), 'Total reportado por Lenovo', '#2ABFDF'),
          kpiCard('Pendiente por asignar', formatInteger(metric('pendiente_asignar')), 'Requiere distribución', '#F0A020', 'program-channel-kpi-alert'),
          kpiCard('Saldo asignado', formatInteger(metric('saldo_asignado')), 'Disponible en el control', '#0DBF82'),
          kpiCard('Plan Ultra Lenovo WS', formatInteger(sedPoints), 'Puntos SED en control separado', '#8B5FC8')
        ].join('');
        return;
      }

      if(!hasDetailFilters && state.filters.canal === 'Lenovo' && unit === 'USD') {
        const movement = name => sumRows(unitRows.filter(row => row.grupo === 'lenovo_card' && row.movimiento === name));
        const summaryValue = (name, fallback) => movement(name) || fallback;
        const available = movement('saldo_disponible') || sumRows(unitRows.filter(row => row.grupo === 'lenovo_card' && row.movimiento === 'saldo_disponible'));
        const topups = summaryValue('resumen_recargas', movement('recarga'));
        const purchases = summaryValue('resumen_compras', movement('compra'));
        const posFees = summaryValue('resumen_comision_pos', movement('comision_pos'));
        const reloadFees = summaryValue('resumen_comision_recarga', movement('comision_recarga'));
        const fees = posFees + reloadFees;
        host.innerHTML = [
          kpiCard('Total recargas', formatMoney(topups, 'USD'), 'Total explícito de la hoja', '#2ABFDF'),
          kpiCard('Compras identificadas', formatMoney(purchases, 'USD'), 'Total explícito de la hoja', '#8B5FC8'),
          kpiCard('Comisiones', formatMoney(fees, 'USD'), `POS ${formatMoney(posFees, 'USD')} · recarga ${formatMoney(reloadFees, 'USD')}`, '#F0A020'),
          kpiCard('Saldo disponible', formatMoney(available, 'USD'), 'Saldo actual de tarjeta', '#0DBF82')
        ].join('');
        return;
      }

      if(!hasDetailFilters && state.filters.canal === 'Lenovo' && unit === 'COP') {
        const ledger = unitRows.filter(row => row.grupo === 'lenovo_card' && row.metrica === 'movement');
        const year2025 = sumRows(unitRows.filter(row => row.movimiento === 'consumo_2025')) || sumRows(ledger.filter(row => row.anio === 2025));
        const year2026 = sumRows(unitRows.filter(row => row.movimiento === 'consumo_2026')) || sumRows(ledger.filter(row => row.anio === 2026));
        const variation = year2025 ? ((year2026 / year2025) - 1) * 100 : 0;
        const balance = sumRows(unitRows.filter(row => row.grupo === 'lenovo_card' && row.movimiento === 'saldo_disponible'));
        host.innerHTML = [
          kpiCard('Consumo 2025', formatMoney(year2025, 'COP'), 'Compras y comisiones registradas', '#2D4FD6'),
          kpiCard('Consumo 2026', formatMoney(year2026, 'COP'), 'Compras y comisiones registradas', '#2ABFDF'),
          kpiCard('Variación 2026 vs. 2025', `${variation >= 0 ? '+' : ''}${variation.toLocaleString('es-CO',{maximumFractionDigits:1})}%`, formatMoney(year2026 - year2025, 'COP'), variation >= 0 ? '#F0A020' : '#0DBF82'),
          kpiCard('Saldo tarjeta', formatMoney(balance, 'COP'), 'Equivalencia reportada', '#0DBF82')
        ].join('');
        return;
      }

      const movements = unitRows.filter(row => row.metrica !== 'balance');
      const balances = unitRows.filter(row => row.metrica === 'balance');
      const pendingRows = unitRows.filter(row => statusGroup(row.estado) === 'pending');
      const completedRows = unitRows.filter(row => ['paid','redeemed'].includes(statusGroup(row.estado)));
      const alertCount = unitRows.filter(row => ['expired','rejected'].includes(statusGroup(row.estado))).length;
      const valueFormatter = value => unit === 'Puntos' ? formatInteger(value) : formatMoney(value, unit);
      host.innerHTML = [
        kpiCard(`${unit === 'Puntos' ? 'Movimientos' : 'Incentivos'} ${unit}`, valueFormatter(sumRows(movements.length ? movements : unitRows)), `${unitRows.length} registros`, '#2ABFDF'),
        kpiCard('Redimidos / entregados', valueFormatter(sumRows(completedRows)), 'Estado completado', '#0DBF82'),
        kpiCard('Disponibles / pendientes', valueFormatter(sumRows(pendingRows) || sumRows(balances)), 'Por gestionar', '#F0A020'),
        kpiCard('Alertas', formatInteger(alertCount), 'Expirados / rechazados', '#E84040', alertCount ? 'program-channel-kpi-alert' : '')
      ].join('');
      return;
    }

    const selectedUnitRows = allUnitsRows.filter(row => row.unidad === state.filters.unidad);

    if(!hasDetailFilters && state.filters.canal === 'Microsoft' && selectedUnitRows.length) {
      const sent2025 = sumRows(selectedUnitRows.filter(row => row.movimiento === 'pago' && row.anio === 2025));
      const sent2026 = sumRows(selectedUnitRows.filter(row => row.movimiento === 'pago' && row.anio === 2026));
      const upcoming = sumRows(selectedUnitRows.filter(row => row.movimiento === 'proximo'));
      const atRisk = sumRows(selectedUnitRows.filter(row => row.movimiento === 'en_riesgo'));
      const variation = sent2025 ? ((sent2026 / sent2025) - 1) * 100 : 0;
      host.innerHTML = [
        kpiCard('Enviado 2026', formatMoney(sent2026, state.filters.unidad), `${variation >= 0 ? '+' : ''}${variation.toLocaleString('es-CO',{maximumFractionDigits:0})}% vs. periodos disponibles 2025`, '#0DBF82'),
        kpiCard('Enviado 2025', formatMoney(sent2025, state.filters.unidad), 'Histórico disponible: ago–oct', '#2D4FD6'),
        kpiCard('Upcoming', formatMoney(upcoming, state.filters.unidad), 'Próximo pago · jul 2026', '#F0A020'),
        kpiCard('Forfeit en proceso', formatMoney(atRisk, state.filters.unidad), 'Monto en riesgo', '#E84040', atRisk ? 'program-channel-kpi-alert' : '')
      ].join('');
      return;
    }

    if(!hasDetailFilters && state.filters.unidad === 'USD' && !state.filters.canal) {
      const lenovoPaid2026 = sumRows(selectedUnitRows.filter(row => row.canal === 'Lenovo' && row.movimiento === 'pago' && /^FY26/.test(row.periodo))) || BUSINESS_METRICS.lenovo.actual2026;
      const microsoftSent2026 = sumRows(selectedUnitRows.filter(row => row.canal === 'Microsoft' && row.movimiento === 'pago' && row.anio === 2026));
      host.innerHTML = [
        kpiCard('Lenovo pagado 2026', formatMoney(lenovoPaid2026, 'USD'), 'Programas de canal', '#8B5FC8'),
        kpiCard('Microsoft enviado 2026', formatMoney(microsoftSent2026, 'USD'), 'Ene–jun según archivo', '#0DBF82'),
        kpiCard('Dell aprobado a pagar', formatMoney(BUSINESS_METRICS.dell.approved, 'USD'), 'FY26–FY27', '#2D4FD6'),
        kpiCard('ASUS cumplimiento', `${BUSINESS_METRICS.asus.compliance}%`, `Alerta · meta ${formatMoney(BUSINESS_METRICS.asus.quota, 'USD')}`, '#E84040', 'program-channel-kpi-alert')
      ].join('');
      return;
    }

    if(!hasDetailFilters && state.filters.unidad === 'USD' && state.filters.canal === 'Lenovo') {
      const yearly = year => sumRows(selectedUnitRows.filter(row => row.movimiento === 'pago' && row.periodo.startsWith(`FY${year}`)));
      const paid2025 = yearly('25') || BUSINESS_METRICS.lenovo.history.FY25;
      const paid2026 = yearly('26') || BUSINESS_METRICS.lenovo.actual2026;
      const variation = paid2025 ? ((paid2026 / paid2025) - 1) * 100 : 0;
      const engage = sumRows(selectedUnitRows.filter(row => /engage platinum/i.test(row.programa) && /^FY26/.test(row.periodo)));
      host.innerHTML = [
        kpiCard('Pagado 2026', formatMoney(paid2026, 'USD'), 'Tablas dinámicas = pagos', '#2ABFDF'),
        kpiCard('Pagado 2025', formatMoney(paid2025, 'USD'), 'Base comparativa', '#2D4FD6'),
        kpiCard('Variación 2026 vs. 2025', `${variation >= 0 ? '+' : ''}${variation.toLocaleString('es-CO',{maximumFractionDigits:1})}%`, '2026 corresponde al periodo cargado', '#F0A020'),
        kpiCard('Engage Platinum 2026', formatMoney(engage, 'USD'), 'Mayor componente pagado', '#8B5FC8')
      ].join('');
      return;
    }

    if(!hasDetailFilters && state.filters.unidad === 'USD' && state.filters.canal === 'Dell') {
      const dellRows = selectedUnitRows;
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
    if(!host) return;
    const byChannel = !state.filters.canal;
    if(title) title.innerHTML = executiveOverview
      ? 'Pagos e incentivos por canal <span>USD</span>'
      : byChannel
        ? `${state.mode === 'Punto' ? 'Incentivos' : 'Rebates'} por canal <span>${escapeHtml(state.filters.unidad)}</span>`
        : `Composición de ${escapeHtml(state.filters.canal)} <span>${escapeHtml(state.filters.unidad)}</span>`;
    const totals = new Map();
    const movementRows = rows.filter(row => row.metrica !== 'balance');
    let chartRows = state.mode === 'Punto' && movementRows.length ? movementRows : rows;
    if(state.filters.canal === 'Lenovo' && state.mode === 'Punto' && state.filters.unidad === 'COP') {
      const annualSummary = rows.filter(row => /^consumo_20\d{2}$/.test(row.movimiento));
      if(annualSummary.length) chartRows = annualSummary;
    }
    chartRows.forEach(row => {
      let key = row.canal;
      if(!byChannel) {
        if(row.canal === 'Microsoft') key = row.estado;
        else if(row.canal === 'Lenovo' && state.mode === 'Punto' && state.filters.unidad === 'COP' && row.anio) key = String(row.anio);
        else key = row.programa;
      }
      totals.set(key, (totals.get(key) || 0) + row.valor);
    });
    let items = [...totals.entries()].map(([name,val]) => ({name,val})).sort((a,b) => b.val - a.val);
    if(executiveOverview) {
      items = [
        { name:'Lenovo', val:sumRows(rows.filter(row => row.canal === 'Lenovo' && row.movimiento === 'pago' && /^FY26/.test(row.periodo))) || BUSINESS_METRICS.lenovo.actual2026 },
        { name:'Dell', val:BUSINESS_METRICS.dell.approved },
        { name:'Microsoft', val:sumRows(rows.filter(row => row.canal === 'Microsoft' && row.movimiento === 'pago' && row.anio === 2026)) },
        { name:'Epson', val:sumRows(rows.filter(row => row.canal === 'Epson')) }
      ].filter(item => item.val > 0);
    }
    if(!items.length) {
      host.innerHTML = '<div class="program-channel-empty">Sin datos para los filtros seleccionados.</div>';
      return;
    }
    if(typeof renderBars === 'function') {
      const palette = typeof COLORS !== 'undefined' ? COLORS : ['#2D4FD6','#8B5FC8','#2ABFDF','#0DBF82'];
      renderBars('program-channel-bars', items, palette, value =>
        state.filters.unidad === 'Puntos' ? formatInteger(value) : formatCompact(value, state.filters.unidad), {
          nameClass: 'w100',
          getOnClick: byChannel ? item => `ProgramChannelModule.setFilter('canal',${JSON.stringify(item.name)})` : null,
          getIsSelected: byChannel ? item => state.filters.canal === item.name : null,
          clickTitle: 'Filtrar por canal',
          tooltipPrefix: 'Filtrar canal: '
        }
      );
    }
  }

  function renderTrend(rows){
    const title = document.getElementById('program-channel-trend-title');
    const host = document.getElementById('program-channel-trend');
    if(title) title.innerHTML = `Evolución por periodo <span>${escapeHtml(state.filters.unidad)}</span>`;
    if(!host) return;
    const totals = new Map();
    const movementRows = rows.filter(row => row.metrica !== 'balance');
    const trendRows = state.mode === 'Punto' && movementRows.length ? movementRows : rows;
    trendRows.filter(row => periodSortValue(row.periodo) >= 0).forEach(row =>
      totals.set(row.periodo, (totals.get(row.periodo) || 0) + row.valor)
    );
    const points = [...totals.entries()]
      .map(([period,val]) => ({period,val}))
      .sort((a,b) => periodSortValue(a.period) - periodSortValue(b.period));
    if(!points.length) {
      host.innerHTML = '<div class="program-channel-empty">No hay periodos trimestrales o mensuales para graficar.</div>';
      return;
    }

    const width = 760, height = 238, left = 128, right = 18, top = 18, bottom = 42;
    const graphWidth = width - left - right;
    const graphHeight = height - top - bottom;
    const maxValue = Math.max(...points.map(point => point.val), 1);
    const xFor = index => points.length === 1 ? left + graphWidth / 2 : left + index * graphWidth / (points.length - 1);
    const yFor = value => top + graphHeight - (value / maxValue) * graphHeight;
    const line = points.map((point,index) => `${xFor(index).toFixed(1)},${yFor(point.val).toFixed(1)}`).join(' ');
    const area = `${left},${top + graphHeight} ${line} ${xFor(points.length - 1)},${top + graphHeight}`;
    const labelStep = Math.max(1, Math.ceil(points.length / 8));
    let svg = `<svg viewBox="0 0 ${width} ${height}" class="program-channel-trend-svg" role="img" aria-label="Evolución por periodo">`;
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
    const byChannel = !state.filters.canal;
    if(title) title.innerHTML = `${byChannel ? 'Ranking de canales' : 'Principales conceptos'} <span>${escapeHtml(state.filters.unidad)}</span>`;
    const executiveOverview = state.mode === 'Rebate' && state.filters.unidad === 'USD' &&
      !state.filters.canal && !state.filters.periodo && !state.filters.estado;
    let items = [];
    if(executiveOverview) {
      items = [
        { name:'Lenovo', val:sumRows(rows.filter(row => row.canal === 'Lenovo' && row.movimiento === 'pago' && /^FY26/.test(row.periodo))) || BUSINESS_METRICS.lenovo.actual2026 },
        { name:'Dell', val:BUSINESS_METRICS.dell.approved },
        { name:'Microsoft', val:sumRows(rows.filter(row => row.canal === 'Microsoft' && row.movimiento === 'pago' && row.anio === 2026)) },
        { name:'Epson', val:sumRows(rows.filter(row => row.canal === 'Epson')) }
      ].filter(item => item.val > 0).sort((a,b) => b.val - a.val).slice(0, 3);
    } else {
      const movementRows = rows.filter(row => row.metrica !== 'balance');
      let rankingRows = state.mode === 'Punto' && movementRows.length ? movementRows : rows;
      if(state.filters.canal === 'Lenovo' && state.mode === 'Punto' && state.filters.unidad === 'COP') {
        const annualSummary = rows.filter(row => /^consumo_20\d{2}$/.test(row.movimiento));
        if(annualSummary.length) rankingRows = annualSummary;
      }
      const totals = new Map();
      rankingRows.forEach(row => {
        const key = byChannel ? row.canal : row.canal === 'Microsoft' ? row.estado : row.programa;
        totals.set(key, (totals.get(key) || 0) + row.valor);
      });
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
      <button type="button" class="program-channel-ranking-row" ${byChannel ? `onclick="ProgramChannelModule.setFilter('canal',${escapeAttr(JSON.stringify(item.name))})" title="Filtrar canal ${escapeAttr(item.name)}"` : 'aria-disabled="true"'}>
        <span class="program-channel-ranking-position">${index + 1}</span>
        <span class="program-channel-ranking-name">${escapeHtml(item.name)}</span>
        <span class="program-channel-ranking-value">${escapeHtml(state.filters.unidad === 'Puntos' ? formatInteger(item.val) : formatCompact(item.val, state.filters.unidad))}</span>
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
        Lenovo: insightCard('info','i','Lenovo','Las tablas dinámicas se presentan como pagos confirmados. 2026 acumula USD 20.633,90 frente a USD 68.165,62 en 2025.'),
        Dell: insightCard('info','i','Dell','FY26 ejecutado y pagado; FY27 mantiene fondos aún no convertidos en ingreso.'),
        ASUS: insightCard('danger','!','ASUS','Solo 19% de cumplimiento de cuota: acción comercial prioritaria.'),
        Epson: insightCard('warning','•','Epson','Separar Consumo y Comercial; monitorear Pendiente vs. Procesado por mayorista.'),
        HPE: insightCard('info','i','HPE','La información validada disponible corresponde a puntos, no a rebates.'),
        Intel: insightCard('info','i','Intel','La información validada disponible corresponde a incentivos en puntos.'),
        Microsoft: insightCard('warning','!','Microsoft','Separar Sent, Upcoming y Forfeit In Progress. El histórico cargado de 2025 solo cubre agosto–octubre.')
      };
      cards = channel
        ? [definitions[channel]].filter(Boolean)
        : [definitions.Lenovo, definitions.Microsoft, definitions.Dell, definitions.ASUS];
    } else {
      const definitions = {
        Dell: insightCard('warning','!','Dell MyRewards','265 puntos pendientes por redimir — acción requerida por gerencia.'),
        HPE: insightCard('warning','!','HPE Instant On','USD 1.905 redimidos; 5 de 9 canjes aún no tienen soporte. La tarjeta USD 300 equivale a COP 990.000 en el Excel.'),
        Lenovo: insightCard('opportunity','★','Lenovo','4.190 puntos emitidos: 3.923 pendientes por asignar y 267 disponibles. Consumo 2026 supera 2025 en 68%.'),
        Intel: insightCard('warning','•','Intel','7.550 puntos pendientes, con equivalencia total de COP 24.915.000.'),
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
    const sheets = uniqueSorted(rows.map(row => row.hoja));
    if(title) title.textContent = `Resumen ejecutivo por canal · ${state.mode === 'Punto' ? 'Puntos e incentivos' : 'Rebates'}`;
    if(meta) meta.textContent = `${rows.length.toLocaleString('es-CO')} registros filtrados · ${sheets.length ? `${sheets.length} hoja(s): ${sheets.join(', ')}` : 'base validada'}`;
    if(exportButton) exportButton.disabled = !rows.length;
    tableHost.innerHTML = `<table class="responsive-table program-channel-table">
      <thead><tr><th>Canal</th><th>Tipo</th><th>Programa</th><th>Fecha</th><th>Periodo</th><th>Valor</th><th>Unidad</th><th>Estado</th><th>Cliente / Ref</th><th>Hoja</th></tr></thead>
      <tbody>${visible.length ? visible.map(row => {
        const group = statusGroup(row.estado);
        return `<tr>
          <td data-label="Canal" style="color:var(--text);font-weight:700">${escapeHtml(row.canal)}</td>
          <td data-label="Tipo"><span class="program-channel-type program-channel-type-${row.tipo.toLowerCase()}">${escapeHtml(row.tipo)}</span></td>
          <td data-label="Programa" class="program-channel-program" title="${escapeAttr(row.programa)}">${escapeHtml(row.programa)}</td>
          <td data-label="Fecha" class="td-mono">${escapeHtml(row.fecha || '—')}</td>
          <td data-label="Periodo" class="td-mono">${escapeHtml(row.periodo)}</td>
          <td data-label="Valor" class="td-mono program-channel-value">${escapeHtml(row.unidad === 'Puntos' ? formatInteger(row.valor) : formatMoney(row.valor, row.unidad))}</td>
          <td data-label="Unidad">${escapeHtml(row.unidad)}</td>
          <td data-label="Estado"><span class="program-channel-state program-channel-state-${group}">${escapeHtml(row.estado)}</span></td>
          <td data-label="Cliente / Ref" class="program-channel-reference" title="${escapeAttr(row.clienteRef)}">${escapeHtml(row.clienteRef)}</td>
          <td data-label="Hoja" title="${escapeAttr(row.hoja)}">${escapeHtml(row.hoja || 'Base validada')}</td>
        </tr>`;
      }).join('') : '<tr><td colspan="10" class="program-channel-empty-cell">Sin registros para los filtros seleccionados.</td></tr>'}</tbody>
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

  function getActiveWorkbookSheet(){
    return state.workbook.sheets.find(sheet => sheet.name === state.workbook.activeSheet) ||
      state.workbook.sheets[0] || null;
  }

  function getWorkbookHeaderIndex(rows){
    const limit = Math.min(rows.length, 8);
    let bestIndex = 0;
    let bestCount = 0;
    for(let index = 0; index < limit; index++) {
      const count = (rows[index] || []).filter(cell => !isEmptyWorkbookCell(cell)).length;
      if(count > bestCount) {
        bestCount = count;
        bestIndex = index;
      }
    }
    return bestIndex;
  }

  function workbookRowMatchesSearch(row, search){
    if(!search) return true;
    const needle = normalizeKey(search);
    return (row || []).some(cell => normalizeKey(cell && cell.text).includes(needle));
  }

  function renderWorkbookSheetTabs(){
    const host = document.getElementById('program-channel-workbook-tabs');
    if(!host) return;
    if(!state.workbook.sheets.length) {
      host.innerHTML = '';
      return;
    }
    host.innerHTML = state.workbook.sheets.map(sheet => {
      const active = sheet.name === state.workbook.activeSheet;
      return `<button type="button" class="program-channel-workbook-tab${active ? ' active' : ''}" role="tab" aria-selected="${active ? 'true' : 'false'}" onclick="ProgramChannelModule.setWorkbookSheet(${escapeAttr(JSON.stringify(sheet.name))})">${escapeHtml(sheet.name)}</button>`;
    }).join('');
  }

  function getWorkbookMergeMaps(sheet){
    const startMap = new Map();
    const covered = new Set();
    (sheet.merges || []).forEach(merge => {
      const rowspan = Math.max(1, merge.endRow - merge.startRow + 1);
      const colspan = Math.max(1, merge.endCol - merge.startCol + 1);
      startMap.set(`${merge.startRow}:${merge.startCol}`, { rowspan, colspan });
      for(let row = merge.startRow; row <= merge.endRow; row++) {
        for(let col = merge.startCol; col <= merge.endCol; col++) {
          if(row === merge.startRow && col === merge.startCol) continue;
          covered.add(`${row}:${col}`);
        }
      }
    });
    return { startMap, covered };
  }

  function firstTextInRow(row){
    const cell = (row || []).find(item => normalizeText(item && item.text));
    return cell ? cell.text : '';
  }

  function getWorkbookSheetTitle(sheet){
    const rows = sheet.rows || [];
    const title = firstTextInRow(rows[0]);
    const subtitle = firstTextInRow(rows[1]);
    return {
      title: title || sheet.name,
      subtitle: subtitle || 'Informe original por hoja'
    };
  }

  function numericCellValue(cell){
    if(cell && typeof cell.raw === 'number' && Number.isFinite(cell.raw)) return cell.raw;
    return null;
  }

  function getBestWorkbookTable(rows){
    let best = { headerIndex: -1, score: 0, numericCols: [], labelCol: 0 };
    const scanLimit = Math.min(rows.length, 18);
    for(let rowIndex = 0; rowIndex < scanLimit; rowIndex++) {
      const row = rows[rowIndex] || [];
      const labelCols = [];
      row.forEach((cell, colIndex) => {
        if(normalizeText(cell && cell.text) && numericCellValue(cell) == null) labelCols.push(colIndex);
      });
      const numericCols = [];
      const lookahead = rows.slice(rowIndex + 1, Math.min(rows.length, rowIndex + 12));
      row.forEach((_, colIndex) => {
        const count = lookahead.filter(nextRow => numericCellValue((nextRow || [])[colIndex]) != null).length;
        if(count >= 2) numericCols.push(colIndex);
      });
      const nonEmpty = row.filter(cell => !isEmptyWorkbookCell(cell)).length;
      const score = numericCols.length * 4 + labelCols.length + nonEmpty;
      if(numericCols.length && score > best.score) {
        best = {
          headerIndex: rowIndex,
          score,
          numericCols,
          labelCol: labelCols[0] != null ? labelCols[0] : 0
        };
      }
    }
    return best.headerIndex >= 0 ? best : null;
  }

  function formatWorkbookMetric(value, unitHint){
    const suffix = normalizeText(unitHint);
    if(/cop/i.test(suffix)) return '$ ' + Math.round(value).toLocaleString('es-CO');
    if(/usd/i.test(suffix)) return 'USD ' + value.toLocaleString('es-CO', { minimumFractionDigits:2, maximumFractionDigits:2 });
    return value.toLocaleString('es-CO', { maximumFractionDigits:2 });
  }

  function buildWorkbookSummaryCards(sheet, tableInfo){
    const rows = sheet.rows || [];
    if(!tableInfo) return [];
    const header = rows[tableInfo.headerIndex] || [];
    const rowsAfter = rows.slice(tableInfo.headerIndex + 1);
    const cards = [];
    tableInfo.numericCols.slice(0, 4).forEach(colIndex => {
      let totalRow = rowsAfter.find(row => /total/i.test(normalizeText((row || [])[tableInfo.labelCol] && row[tableInfo.labelCol].text)) && numericCellValue((row || [])[colIndex]) != null);
      const value = totalRow ? numericCellValue(totalRow[colIndex]) : rowsAfter.reduce((sum, row) => {
        const value = numericCellValue((row || [])[colIndex]);
        return value == null ? sum : sum + value;
      }, 0);
      const label = normalizeText(header[colIndex] && header[colIndex].text) || XLSX.utils.encode_col(colIndex);
      if(value) cards.push({ label, value, display: formatWorkbookMetric(value, label) });
    });
    return cards;
  }

  function buildWorkbookChart(sheet, tableInfo){
    const rows = sheet.rows || [];
    if(!tableInfo) return null;
    const header = rows[tableInfo.headerIndex] || [];
    const unitFor = colIndex => {
      const label = normalizeText(header[colIndex] && header[colIndex].text);
      if(/cop/i.test(label)) return 'COP';
      if(/usd/i.test(label)) return 'USD';
      if(/puntos|points/i.test(label)) return 'Puntos';
      return 'Otro';
    };
    const grouped = tableInfo.numericCols.reduce((acc, colIndex) => {
      const unit = unitFor(colIndex);
      if(!acc[unit]) acc[unit] = [];
      acc[unit].push(colIndex);
      return acc;
    }, {});
    const preferredUnit = grouped.COP ? 'COP' : grouped.USD ? 'USD' : grouped.Puntos ? 'Puntos' : 'Otro';
    const valueCols = (grouped[preferredUnit] || tableInfo.numericCols).slice(0, 2);
    const dataRows = rows.slice(tableInfo.headerIndex + 1)
      .map(row => {
        const label = normalizeText(row[tableInfo.labelCol] && row[tableInfo.labelCol].text);
        const values = valueCols.map(colIndex => numericCellValue(row[colIndex]) || 0);
        return { label, values };
      })
      .filter(item => item.label && !/^total$/i.test(item.label) && item.values.some(value => value !== 0))
      .slice(0, 10);
    if(!dataRows.length || !valueCols.length) return null;
    const max = Math.max(...dataRows.flatMap(item => item.values.map(Math.abs)), 1);
    const series = valueCols.map(colIndex => normalizeText(header[colIndex] && header[colIndex].text) || XLSX.utils.encode_col(colIndex));
    return { dataRows, max, series };
  }

  function renderWorkbookOverview(sheet){
    const host = document.getElementById('program-channel-workbook-overview');
    if(!host) return;
    if(!sheet || state.workbook.status !== 'loaded') {
      host.innerHTML = '';
      return;
    }
    const { title, subtitle } = getWorkbookSheetTitle(sheet);
    const tableInfo = getBestWorkbookTable(sheet.rows || []);
    const cards = buildWorkbookSummaryCards(sheet, tableInfo);
    const chart = buildWorkbookChart(sheet, tableInfo);
    const cardsHtml = cards.length ? `<div class="program-channel-report-kpis">${cards.map(card => `
      <article class="program-channel-report-kpi">
        <span>${escapeHtml(card.label)}</span>
        <strong>${escapeHtml(card.display)}</strong>
      </article>`).join('')}</div>` : '';
    const chartHtml = chart ? `<div class="program-channel-report-chart">
      <div class="program-channel-report-chart-title">${escapeHtml(chart.series.join(' / '))}</div>
      ${chart.dataRows.map(item => `<div class="program-channel-report-bar-row">
        <div class="program-channel-report-bar-label" title="${escapeAttr(item.label)}">${escapeHtml(item.label)}</div>
        <div class="program-channel-report-bar-stack">
          ${item.values.map((value, index) => `<div class="program-channel-report-bar-wrap">
            <span class="program-channel-report-bar series-${index}" style="width:${Math.max(2, Math.min(100, Math.abs(value) / chart.max * 100)).toFixed(2)}%"></span>
            <em>${escapeHtml(formatWorkbookMetric(value, chart.series[index]))}</em>
          </div>`).join('')}
        </div>
      </div>`).join('')}
    </div>` : '<div class="program-channel-report-empty">Esta hoja no tiene una serie numérica suficiente para graficar; se muestra la tabla original completa.</div>';
    host.innerHTML = `<section class="program-channel-report-hero">
      <div>
        <div class="program-channel-report-eyebrow">Hoja del informe</div>
        <h3>${escapeHtml(title)}</h3>
        <p>${escapeHtml(subtitle)}</p>
      </div>
      ${cardsHtml}
    </section>${chartHtml}`;
  }

  function renderWorkbookReport(){
    const sheetSelect = document.getElementById('program-channel-workbook-sheet');
    const searchInput = document.getElementById('program-channel-workbook-search');
    const meta = document.getElementById('program-channel-workbook-meta');
    const status = document.getElementById('program-channel-workbook-status');
    const tableHost = document.getElementById('program-channel-workbook-table');
    if(!tableHost) return;

    if(sheetSelect) {
      sheetSelect.innerHTML = state.workbook.sheets.map(sheet =>
        `<option value="${escapeAttr(sheet.name)}"${sheet.name === state.workbook.activeSheet ? ' selected' : ''}>${escapeHtml(sheet.name)}</option>`
      ).join('');
      sheetSelect.disabled = !state.workbook.sheets.length;
    }
    renderWorkbookSheetTabs();
    if(searchInput && searchInput.value !== state.workbook.search) searchInput.value = state.workbook.search;

    const activeSheet = getActiveWorkbookSheet();
    const loadedCount = state.workbook.sheets.length;
    if(meta) {
      meta.textContent = `${state.workbook.fileName || EMBEDDED_REPORT_FILE} · ${loadedCount ? `${loadedCount} hojas cargadas` : 'pendiente de carga'}`;
    }
    if(state.workbook.status === 'idle') {
      if(status) status.textContent = 'Cargando el informe incluido en el proyecto...';
      renderWorkbookOverview(null);
      tableHost.innerHTML = '<div class="program-channel-empty">Preparando informe Excel original.</div>';
      return;
    }
    if(state.workbook.status === 'loading') {
      if(status) status.textContent = 'Leyendo informe Excel original...';
      renderWorkbookOverview(null);
      tableHost.innerHTML = '<div class="program-channel-empty">Leyendo hojas y valores guardados.</div>';
      return;
    }
    if(state.workbook.status === 'error') {
      if(status) status.textContent = state.workbook.error;
      renderWorkbookOverview(null);
      tableHost.innerHTML = '<div class="program-channel-empty">No se pudo cargar automaticamente el archivo. Usa Cargar Excel para seleccionarlo.</div>';
      return;
    }
    if(!activeSheet) {
      if(status) status.textContent = 'El informe no contiene hojas visibles.';
      tableHost.innerHTML = '<div class="program-channel-empty">Sin hojas para mostrar.</div>';
      renderWorkbookOverview(null);
      return;
    }

    renderWorkbookOverview(activeSheet);
    const rows = activeSheet.rows || [];
    const mergeMaps = getWorkbookMergeMaps(activeSheet);
    const bodyRows = rows
      .map((row, index) => ({ row, number: index + 1 }))
      .filter(item => workbookRowMatchesSearch(item.row, state.workbook.search));
    const visibleRows = bodyRows.slice(0, 160);
    const colCount = Math.max(...rows.map(row => row.length), ...visibleRows.map(item => item.row.length), 1);
    if(status) {
      const filtered = state.workbook.search ? `${bodyRows.length.toLocaleString('es-CO')} filas filtradas` : `${rows.length.toLocaleString('es-CO')} filas`;
      status.textContent = `${activeSheet.name} · ${filtered} · valores mostrados completos`;
    }

    const headerHtml = '<th class="program-channel-workbook-row-head">#</th>' +
      Array.from({ length: colCount }, (_, index) =>
        `<th>${escapeHtml(XLSX.utils.encode_col(index))}</th>`
      ).join('');
    const rowsHtml = visibleRows.map(item => {
      const rowIndex = item.number - 1;
      let cellsHtml = '';
      for(let index = 0; index < colCount; index++) {
        const mergeKey = `${rowIndex}:${index}`;
        if(mergeMaps.covered.has(mergeKey)) continue;
        const cell = item.row[index] || { text:'' };
        const merge = mergeMaps.startMap.get(mergeKey);
        const spanAttrs = merge
          ? `${merge.rowspan > 1 ? ` rowspan="${merge.rowspan}"` : ''}${merge.colspan > 1 ? ` colspan="${merge.colspan}"` : ''}`
          : '';
      const isNumber = typeof cell.raw === 'number';
        cellsHtml += `<td${spanAttrs} class="${isNumber ? 'td-mono program-channel-workbook-number' : ''}" title="${escapeAttr(cell.formula ? '=' + cell.formula : cell.text)}">${escapeHtml(cell.text)}</td>`;
      }
      return `<tr><td class="td-mono program-channel-workbook-row-head">${item.number}</td>${cellsHtml}</tr>`;
    }).join('');
    tableHost.innerHTML = `<table class="responsive-table program-channel-workbook-table">
      <thead><tr>${headerHtml}</tr></thead>
      <tbody>${rowsHtml || `<tr><td colspan="${colCount + 1}" class="program-channel-empty-cell">Sin filas para mostrar.</td></tr>`}</tbody>
    </table>`;
  }

  function render(){
    const page = document.getElementById('page-programas');
    if(!page || !canAccess()) return;
    if(state.workbook.status === 'idle') reloadEmbeddedReport();
    renderWorkbookReport();
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
        { header:'Tipo(Rebate/Punto-Incentivo)', key:'tipo', width:30 },
        { header:'Programa', key:'programa', width:52 },
        { header:'Fecha', key:'fecha', width:14 },
        { header:'Periodo(FY-Q/Mes)', key:'periodo', width:18 },
        { header:'Valor', key:'valor', width:18 },
        { header:'Unidad(USD/COP/Puntos)', key:'unidad', width:24 },
        { header:'Estado', key:'estado', width:24 },
        { header:'Cliente/Ref', key:'clienteRef', width:42 },
        { header:'Movimiento', key:'movimiento', width:20 },
        { header:'Hoja origen', key:'hoja', width:30 }
      ];
      rows.forEach(row => worksheet.addRow({
        canal: row.canal,
        tipo: row.tipo,
        programa: row.programa,
        fecha: row.fecha,
        periodo: row.periodo,
        valor: row.valor,
        unidad: row.unidad,
        estado: row.estado,
        clienteRef: row.clienteRef,
        movimiento: row.movimiento,
        hoja: row.hoja || 'Base validada'
      }));
      worksheet.autoFilter = { from:'A1', to:'K1' };
      worksheet.getRow(1).height = 26;
      worksheet.getRow(1).font = { name:'Aptos Display', size:10, bold:true, color:{ argb:'FFFFFFFF' } };
      worksheet.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0B1956' } };
      worksheet.getRow(1).alignment = { vertical:'middle', horizontal:'center' };
      for(let index = 2; index <= rows.length + 1; index++) {
        const row = worksheet.getRow(index);
        row.font = { name:'Aptos', size:9, color:{ argb:'FF172033' } };
        row.alignment = { vertical:'middle' };
        row.getCell(3).alignment = { vertical:'middle', wrapText:true };
        row.getCell(6).numFmt = rows[index - 2].unidad === 'Puntos' ? '#,##0' : '#,##0.00';
        row.getCell(6).alignment = { horizontal:'right' };
        row.getCell(9).alignment = { vertical:'middle', wrapText:true };
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
    reloadEmbeddedReport,
    openLocalFiles,
    handleLocalFiles,
    setMode,
    setFilter,
    clearFilters,
    setPage,
    setWorkbookSheet,
    setWorkbookSearch,
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
