// ══════════════════════════════════════════════════════════════════════
//   MÓDULO: FONDOS DE MERCADEO (MDF) - FORECAST 2026
// ══════════════════════════════════════════════════════════════════════

(function () {
  'use strict';

  // Usuarios estrictamente autorizados para acceder a esta vista
  const ALLOWED_USERS = [
    'nini.beltran@provexpress.com.co',
    'c.estrategica@provexpress.com.co',
    'juannovoa@provexpress.com.co',
    'especialista.preventa@provexpress.com.co'
  ];

  // Datos finales validados del libro FONDOS DE MERCADEO NINI_final.xlsx.
  const BASELINE_DATA = {
    trmDefault: 3500,
    lastUpdate: '2026-08-14',
    sourceFileName: 'FONDOS DE MERCADEO NINI_final.xlsx',
    brands: [
      {
        id: 'hp',
        name: 'HP',
        shortName: 'Hp',
        color: '#0096D6',
        bgAlpha: 'rgba(0, 150, 214, 0.18)',
        incomes: [
          { period: '26Q1', usd: 2200, trm: 3861, cop: 8494200, obs: 'Factura 108962' },
          { period: '26Q1', usd: 500, trm: 4019, cop: 2009500, obs: 'Factura 110438' },
          { period: '26Q1', usd: 1500, trm: 3861, cop: 5791500, obs: 'Factura 106826' },
          { period: '26Q1', usd: 1500, trm: 3861, cop: 5791500, obs: 'Factura 107237' },
          { period: '26Q2', usd: 3000, trm: 3668, cop: 11004000, obs: 'Factura 112747' },
          { period: '26Q2', usd: 3000, trm: 3668, cop: 11004000, obs: 'Factura 113461' },
          { period: '26Q2', usd: 1000, trm: 3861, cop: 3861000, obs: 'Factura 106957' },
          { period: '26Q3', usd: 3000, trm: 3668, cop: 11004000, obs: 'Factura 113701' },
          { period: '26Q3', usd: 1000, trm: 3668, cop: 3668000, obs: 'Factura 112763' }
        ],
        outflows: [
          { type: 'Evento', concept: 'HP Suministros - Evento', cupos: null, cop: 2710400, obs: 'Evento presencial suministros' },
          { type: 'Evento', concept: 'Liga Z - Capacitación', cupos: null, cop: 486000, obs: 'Capacitación fuerza de ventas' },
          { type: 'Evento', concept: 'HP Workstation - Hotel Hilton', cupos: null, cop: 730085, obs: 'Evento Workstation' },
          { type: 'Evento', concept: 'HP Nueva era de la tecnología - Hotel', cupos: null, cop: 5943183, obs: 'Convención hotelera' },
          { type: 'Incentivo', concept: 'Concurso Supplies', cupos: null, cop: 2500000, obs: 'Premios canal suministros' },
          { type: 'Incentivo', concept: 'Concurso venta portafolio HP', cupos: null, cop: 3000000, obs: 'Incentivo vendedores' },
          { type: 'Incentivo', concept: 'Campaña Suministros', cupos: null, cop: 6000000, obs: 'Campaña comercial' },
          { type: 'Incentivo', concept: 'Camisetas (obsequio gerencia comercial y direcciones)', cupos: null, cop: 1500000, obs: 'Dotación corporativa' },
          { type: 'Incentivo', concept: 'Incentivo Suministros (Bono Adidas ganadora Dilma)', cupos: null, cop: 300000, obs: 'Bono regalo' },
          { type: 'Incentivo', concept: 'Nintendo', cupos: null, cop: 2600000, obs: 'Premio meta de ventas' },
          { type: 'Rio', concept: 'Río (20 al 24 mayo) - cupos @ $6.500.000', cupos: 5, cop: 32500000, obs: '5 cupos según consolidado final' }
        ],
        notes: 'Informe final: ingresos $62.627.700, salidas $58.269.668 y saldo $4.358.032.'
      },
      {
        id: 'dell',
        name: 'DELL',
        shortName: 'Dell',
        color: '#0076CE',
        bgAlpha: 'rgba(0, 118, 206, 0.18)',
        incomes: [
          { period: '2026 Q3', usd: 1414.49, trm: 3513.54, cop: 4969868, obs: 'Agosto - Octubre · factura 114027' },
          { period: '2026 Q4', usd: 1000, trm: 3810.99, cop: 3810990, obs: 'Noviembre 1 - enero 30 · factura 104032' },
          { period: '2027 Q1', usd: 2864, trm: 3066.99, cop: 8783850, obs: 'Enero 31 - mayo 1 · factura 114319' },
          { period: '2027 Q2', usd: 1811, trm: 3962.55, cop: 7176180, obs: 'Mayo 2 - julio 31 · factura 114027' },
          { period: 'Proposal ingresado', usd: 6500, trm: 3500, cop: 22750000, obs: 'Fondos Proposal ingresados por NC' },
          { period: 'Proposal pendiente', usd: 3000, trm: 3500, cop: 10500000, obs: 'Incluido en el total final; pendiente por ingresar' }
        ],
        outflows: [
          { type: 'Evento', concept: 'Infraestructura', cupos: null, cop: 8874800, obs: '24 de abril 2026' },
          { type: 'Evento', concept: 'Tecnología cómputo', cupos: null, cop: 6452650, obs: '17 de julio 2026' },
          { type: 'Incentivo', concept: 'Televisor, Apple TV, bonos H&M', cupos: null, cop: 3000000, obs: 'Premios fidelización' },
          { type: 'Incentivo', concept: 'Play Station', cupos: null, cop: 2600000, obs: 'Premio mejor ejecutivo' },
          { type: 'Rio', concept: 'Río (20 al 24 mayo) - cupos @ $6.500.000', cupos: 5, cop: 32500000, obs: '5 cupos según consolidado final' }
        ],
        notes: 'Informe final: el total incluye $10.500.000 de Proposal pendientes por ingresar; saldo reportado $4.563.438.'
      },
      {
        id: 'lenovo',
        name: 'LENOVO',
        shortName: 'Lenovo',
        color: '#E2231A',
        bgAlpha: 'rgba(226, 35, 26, 0.18)',
        incomes: [
          { period: 'Validar NC', usd: 3349, trm: 3500, cop: 11721500, obs: 'Validar nota crédito' },
          { period: 'Ingreso', usd: 1906, trm: 3500, cop: 6671000, obs: 'Ingreso confirmado en informe final' }
        ],
        outflows: [
          { type: 'Evento', concept: 'Infraestructura', cupos: null, cop: 4408950, obs: 'Evento técnico Lenovo' },
          { type: 'Incentivo', concept: 'Lenovo Legion Go S', cupos: null, cop: 2500000, obs: 'Incentivo comercial' },
          { type: 'Incentivo', concept: 'Bono María Paola Briceño', cupos: null, cop: 500000, obs: 'Reconocimiento especial' },
          { type: 'Rio', concept: 'Rio (20 al 24 mayo) - cupos @ $6.500.000', cupos: 1.5, cop: 9750000, obs: '1.5 cupos convención Río' }
        ],
        notes: 'Informe final: ingresos $18.392.500, salidas $17.158.950 y saldo $1.233.550.'
      },
      {
        id: 'intel',
        name: 'INTEL',
        shortName: 'Intel',
        color: '#0071C5',
        bgAlpha: 'rgba(0, 113, 197, 0.18)',
        incomes: [
          { period: 'H1', usd: 4990, trm: 3500, cop: 17465000, obs: 'Asignación semestral H1' }
        ],
        outflows: [
          { type: 'Evento', concept: 'Evento (USD 490)', cupos: null, cop: 1715000, obs: 'Evento focalizado' },
          { type: 'Incentivo', concept: 'Capacitación Rosario - código de vestimenta', cupos: null, cop: 6000000, obs: 'Capacitación e imagen corporativa' },
          { type: 'Rio', concept: 'Rio (20 al 24 mayo) - cupos @ $6.500.000', cupos: 1.5, cop: 9750000, obs: '1.5 cupos convención Río' }
        ],
        notes: 'Informe final: ejecución del 100% y saldo $0.'
      },
      {
        id: 'microsoft',
        name: 'MICROSOFT',
        shortName: 'Microsoft',
        color: '#107C41',
        bgAlpha: 'rgba(16, 124, 65, 0.18)',
        incomes: [
          { period: 'LOL', usd: 3000, trm: 3500, cop: 10500000, obs: 'Fondo mayorista LOL (TRM 3.500)' },
          { period: 'INGRAM', usd: 4897, trm: 3500, cop: 17139500, obs: 'Fondo mayorista Ingram Micro (TRM 3.500)' }
        ],
        outflows: [
          { type: 'Incentivo', concept: 'Cruce de cuentas Licencias', cupos: null, cop: 17139500, obs: 'Compras' },
          { type: 'Rio', concept: 'Río (20 al 24 mayo) - cupos @ $6.500.000', cupos: 1.5, cop: 9750000, obs: 'Pendiente por ingresar según informe final' }
        ],
        notes: 'Informe final: ingresos $27.639.500, salidas $26.889.500 y saldo $750.000.'
      },
      {
        id: 'cisco',
        name: 'CISCO',
        shortName: 'Cisco',
        color: '#1BA0D7',
        bgAlpha: 'rgba(27, 160, 215, 0.18)',
        incomes: [
          { period: 'Único', usd: 3000, trm: 3333.3333333333335, cop: 10000000, obs: 'TRM implícita $3.333,33 (valor fijo COP)' }
        ],
        outflows: [
          { type: 'Rio', concept: 'Rio (20 al 24 mayo) - cupos @ $6.500.000', cupos: 1.5, cop: 9750000, obs: '1.5 cupos convención Río' },
          { type: 'Otro', concept: 'Saldo cruza negocio María Paola', cupos: null, cop: 250000, obs: 'Ajuste interno negocio' }
        ],
        notes: 'Informe final: ejecución del 100% y saldo $0.'
      },
      {
        id: 'champion',
        name: 'CHAMPION',
        shortName: 'Champion',
        color: '#F59E0B',
        bgAlpha: 'rgba(245, 158, 11, 0.18)',
        excludeFromConsolidated: true,
        incomes: [],
        outflows: [
          { brand: 'HP', period: 'Q1', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion HP Q1' },
          { brand: 'HP', period: 'Q2', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion HP Q2' },
          { brand: 'DELL', period: 'Q1', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion DELL Q1' },
          { brand: 'DELL', period: 'Q2', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion DELL Q2' }
        ],
        notes: 'Ejecuciones Champion por $28.000.000. Se muestran para control, pero el informe final no las incluye en el total consolidado de fondos.'
      }
    ]
  };

  // Estado reactivo del módulo
  const state = {
    activeBrandId: 'all',
    activeCategory: 'all',
    data: null,
    charts: {
      barChart: null,
      doughnutChart: null
    }
  };

  function getWorkbookSheet(workbook, name) {
    const target = String(name || '').trim().toLowerCase();
    const sheetName = (workbook.SheetNames || []).find(item => String(item).trim().toLowerCase() === target);
    return sheetName ? workbook.Sheets[sheetName] : null;
  }

  function getCellNumber(sheet, address) {
    const value = sheet && sheet[address] ? sheet[address].v : null;
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const normalized = String(value == null ? '' : value)
      .replace(/[^0-9,.-]/g, '')
      .replace(/\.(?=\d{3}(?:\D|$))/g, '')
      .replace(',', '.');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function getCellText(sheet, address, fallback = '') {
    const cell = sheet && sheet[address];
    const value = cell ? (cell.w != null ? cell.w : cell.v) : '';
    return String(value == null ? '' : value).trim() || fallback;
  }

  function replaceBrandRecords(data, brandId, incomes, outflows, notes) {
    const brand = data.brands.find(item => item.id === brandId);
    if (!brand) return;
    brand.incomes = incomes.filter(item => item.cop || item.usd);
    brand.outflows = outflows.filter(item => item.cop);
    if (notes) brand.notes = notes;
  }

  function parseFinalWorkbook(workbook, fileName, lastModified) {
    const data = JSON.parse(JSON.stringify(BASELINE_DATA));
    data.sourceFileName = fileName || BASELINE_DATA.sourceFileName;
    data.lastUpdate = lastModified ? String(lastModified).slice(0, 10) : BASELINE_DATA.lastUpdate;

    const hp = getWorkbookSheet(workbook, 'hp');
    const dell = getWorkbookSheet(workbook, 'Dell');
    const lenovo = getWorkbookSheet(workbook, 'Lenovo');
    const intel = getWorkbookSheet(workbook, 'Intel');
    const microsoft = getWorkbookSheet(workbook, 'Microsoft');
    const cisco = getWorkbookSheet(workbook, 'Cisco');
    const champion = getWorkbookSheet(workbook, 'Champion');
    const totalSheet = getWorkbookSheet(workbook, 'Total Fondos');
    if (!hp || !dell || !lenovo || !intel || !microsoft || !cisco || !totalSheet) {
      throw new Error('El informe no contiene todas las hojas esperadas de Fondos Marketing.');
    }

    replaceBrandRecords(data, 'hp',
      Array.from({ length: 9 }, (_, index) => {
        const row = index + 2;
        return {
          period: getCellText(hp, `E${row}`, `Registro ${index + 1}`),
          usd: getCellNumber(hp, `B${row}`),
          trm: getCellNumber(hp, `C${row}`),
          cop: getCellNumber(hp, `D${row}`),
          obs: `Factura ${getCellText(hp, `F${row}`, '-')}`
        };
      }),
      [
        ...Array.from({ length: 4 }, (_, index) => {
          const row = index + 2;
          return { type: 'Evento', concept: getCellText(hp, `I${row}`, 'Evento HP'), cop: getCellNumber(hp, `H${row}`), obs: 'Informe final' };
        }),
        ...Array.from({ length: 6 }, (_, index) => {
          const row = index + 2;
          return { type: 'Incentivo', concept: getCellText(hp, `K${row}`, 'Incentivo HP'), cop: getCellNumber(hp, `J${row}`), obs: 'Informe final' };
        }),
        { type: 'Rio', concept: 'Río (20 al 24 mayo)', cupos: getCellNumber(hp, 'L2') / 6500000 || 5, cop: getCellNumber(hp, 'L2'), obs: getCellText(hp, 'M2', 'Informe final') }
      ],
      `Informe final: ingresos ${formatCOP(getCellNumber(hp, 'D11'))}, salidas ${formatCOP(getCellNumber(hp, 'H12'))} y saldo ${formatCOP(getCellNumber(hp, 'D12'))}.`
    );

    replaceBrandRecords(data, 'dell',
      [
        ...Array.from({ length: 4 }, (_, index) => {
          const row = index + 2;
          return {
            period: `${getCellText(dell, `C${row}`)} ${getCellText(dell, `D${row}`)}`.trim(),
            usd: getCellNumber(dell, `F${row}`),
            trm: getCellNumber(dell, `I${row}`) || getCellNumber(dell, `G${row}`),
            cop: getCellNumber(dell, `H${row}`),
            obs: `${getCellText(dell, `E${row}`)} · factura ${getCellText(dell, `J${row}`, '-')}`
          };
        }),
        ...[8, 9].map(row => ({
          period: getCellText(dell, `C${row}`, 'Fondos Proposal'),
          usd: getCellNumber(dell, `F${row}`),
          trm: getCellNumber(dell, `G${row}`),
          cop: getCellNumber(dell, `H${row}`),
          obs: getCellText(dell, `I${row}`, 'Informe final')
        }))
      ],
      [
        ...[2, 3].map(row => ({ type: 'Evento', concept: getCellText(dell, `M${row}`, 'Evento Dell'), cop: getCellNumber(dell, `L${row}`), obs: getCellText(dell, `N${row}`, 'Informe final') })),
        ...[2, 3].map(row => ({ type: 'Incentivo', concept: getCellText(dell, `P${row}`, 'Incentivo Dell'), cop: getCellNumber(dell, `O${row}`), obs: 'Informe final' })),
        { type: 'Rio', concept: 'Río (20 al 24 mayo)', cupos: 5, cop: getCellNumber(dell, 'Q2'), obs: 'Informe final' }
      ],
      `Informe final: ingresos ${formatCOP(getCellNumber(dell, 'H11'))}, salidas ${formatCOP(getCellNumber(dell, 'S6'))} y saldo ${formatCOP(getCellNumber(dell, 'H12'))}.`
    );

    replaceBrandRecords(data, 'lenovo',
      [2, 4].map(row => ({ period: getCellText(lenovo, `D${row}`, row === 2 ? 'Validar NC' : 'Ingreso'), usd: getCellNumber(lenovo, `A${row}`), trm: getCellNumber(lenovo, `B${row}`), cop: getCellNumber(lenovo, `C${row}`), obs: 'Informe final' })),
      [
        { type: 'Evento', concept: getCellText(lenovo, 'I2', 'Infraestructura'), cop: getCellNumber(lenovo, 'H2'), obs: 'Informe final' },
        ...[2, 3].map(row => ({ type: 'Incentivo', concept: getCellText(lenovo, `K${row}`, 'Incentivo Lenovo'), cop: getCellNumber(lenovo, `J${row}`), obs: 'Informe final' })),
        { type: 'Rio', concept: 'Río (20 al 24 mayo)', cupos: 1.5, cop: getCellNumber(lenovo, 'L2'), obs: 'Informe final' }
      ],
      `Informe final: ingresos ${formatCOP(getCellNumber(lenovo, 'C5'))}, salidas ${formatCOP(getCellNumber(lenovo, 'M2'))} y saldo ${formatCOP(getCellNumber(lenovo, 'C6'))}.`
    );

    replaceBrandRecords(data, 'intel',
      [{ period: getCellText(intel, 'B2', 'H1'), usd: getCellNumber(intel, 'C2'), trm: getCellNumber(intel, 'D2'), cop: getCellNumber(intel, 'E2'), obs: 'Informe final' }],
      [
        { type: 'Evento', concept: getCellText(intel, 'H2', 'Evento Intel'), cop: getCellNumber(intel, 'G2'), obs: 'Informe final' },
        { type: 'Incentivo', concept: getCellText(intel, 'J2', 'Incentivo Intel'), cop: getCellNumber(intel, 'I2'), obs: 'Informe final' },
        { type: 'Rio', concept: 'Río (20 al 24 mayo)', cupos: 1.5, cop: getCellNumber(intel, 'K2'), obs: 'Informe final' }
      ]
    );

    replaceBrandRecords(data, 'microsoft',
      [2, 3].map(row => ({ period: getCellText(microsoft, `D${row}`, 'Microsoft'), usd: getCellNumber(microsoft, `E${row}`), trm: getCellNumber(microsoft, `F${row}`), cop: getCellNumber(microsoft, `G${row}`), obs: 'Informe final' })),
      [
        { type: 'Incentivo', concept: getCellText(microsoft, 'N2', 'Cruce de cuentas Licencias'), cop: getCellNumber(microsoft, 'M2'), obs: 'Compras' },
        { type: 'Rio', concept: 'Río (20 al 24 mayo)', cupos: 1.5, cop: getCellNumber(microsoft, 'O2'), obs: 'Pendiente por ingresar' }
      ]
    );

    replaceBrandRecords(data, 'cisco',
      [{ period: 'Único', usd: getCellNumber(cisco, 'B2'), trm: getCellNumber(cisco, 'C2'), cop: getCellNumber(cisco, 'D2'), obs: 'Informe final' }],
      [
        { type: 'Rio', concept: 'Río (20 al 24 mayo)', cupos: 1.5, cop: getCellNumber(cisco, 'G2'), obs: 'Informe final' },
        { type: 'Otro', concept: getCellText(cisco, 'H1', 'Saldo cruza negocio'), cop: getCellNumber(cisco, 'H2'), obs: 'Informe final' }
      ]
    );

    if (champion) {
      replaceBrandRecords(data, 'champion', [], [
        ...[2, 3].map(row => ({ brand: 'HP', period: getCellText(champion, `D${row}`), usd: getCellNumber(champion, `E${row}`), trm: 3500, cop: getCellNumber(champion, `F${row}`), type: 'Otro', concept: `Ejecución Champion HP ${getCellText(champion, `D${row}`)}` })),
        ...[2, 3].map(row => ({ brand: 'DELL', period: getCellText(champion, `D${row}`), usd: getCellNumber(champion, `G${row}`), trm: 3500, cop: getCellNumber(champion, `H${row}`), type: 'Otro', concept: `Ejecución Champion DELL ${getCellText(champion, `D${row}`)}` }))
      ]);
    }

    const summary = getGlobalSummary(data.brands.map(getProcessedBrandData));
    const expectedIncome = getCellNumber(totalSheet, 'H11');
    const expectedOutflow = getCellNumber(totalSheet, 'I11');
    const expectedBalance = getCellNumber(totalSheet, 'J11');
    if (Math.abs(summary.totalIncome - expectedIncome) > 1 || Math.abs(summary.totalOutflow - expectedOutflow) > 1 || Math.abs(summary.totalBalance - expectedBalance) > 1) {
      throw new Error('Los totales leídos no coinciden con la hoja Total Fondos.');
    }
    return data;
  }

  async function loadFromSharePoint(siteId, token) {
    if (!canAccess()) return '';
    if (typeof XLSX === 'undefined') throw new Error('La librería XLSX no está disponible.');
    const forecastBase = typeof getForecastBasePath === 'function'
      ? await getForecastBasePath(siteId, token)
      : 'COMERCIAL/FORECAST 2026';
    const folderPath = typeof joinGraphPath === 'function'
      ? joinGraphPath(forecastBase, 'NINI')
      : `${forecastBase}/NINI`;
    if (typeof buildGraphRootUrl !== 'function') throw new Error('No está disponible la conexión con SharePoint.');
    const response = await fetch(buildGraphRootUrl(siteId, folderPath, 'children?$top=100'), {
      headers: { Authorization: `Bearer ${token}` }
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error((payload.error && payload.error.message) || `No se pudo consultar ${folderPath}.`);
    const candidates = (payload.value || [])
      .filter(item => item && item.name && !item.name.startsWith('~$') && /fondos\s+de\s+mercadeo\s+nini.*\.xlsx$/i.test(item.name))
      .sort((a, b) => new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0));
    if (!candidates.length) throw new Error(`No se encontró el informe de fondos en ${folderPath}.`);
    const selected = typeof ensureDriveItemDownloadUrl === 'function'
      ? await ensureDriveItemDownloadUrl(candidates[0], token)
      : candidates[0];
    if (!selected['@microsoft.graph.downloadUrl']) throw new Error('El informe no tiene URL de descarga.');
    const fileResponse = await fetch(selected['@microsoft.graph.downloadUrl']);
    if (!fileResponse.ok) throw new Error(`No se pudo descargar el informe (${fileResponse.status}).`);
    const workbook = XLSX.read(await fileResponse.arrayBuffer(), { type: 'array', cellDates: true });
    state.data = parseFinalWorkbook(workbook, selected.name, selected.lastModifiedDateTime);
    console.info('[FONDOS] informe SharePoint cargado', selected.name);
    if (document.getElementById('page-fondos')?.classList.contains('active')) init();
    return selected.name;
  }

  // ── Permisos de Acceso ──────────────────────────────────────────────
  function canAccess() {
    if (!window.CURRENT_USER || !CURRENT_USER.role) return true; // Vista local / preview
    return CURRENT_USER.role === 'gerencia'; // Exclusivo para Gerencia (Directores excluidos)
  }

  // ── Formateadores ──────────────────────────────────────────────────
  function formatCOP(amount) {
    if (amount == null || isNaN(amount)) return '$ 0';
    const isNeg = amount < 0;
    const absVal = Math.abs(amount);
    const formatted = new Intl.NumberFormat('es-CO', {
      style: 'currency',
      currency: 'COP',
      maximumFractionDigits: 0
    }).format(absVal);
    return isNeg ? `-${formatted}` : formatted;
  }

  function formatUSD(amount) {
    if (amount == null || isNaN(amount)) return 'USD 0';
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      maximumFractionDigits: 2
    }).format(amount);
  }

  function escapeMarkup(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatCompactCOP(amount) {
    const value = Number(amount || 0);
    const abs = Math.abs(value);
    const sign = value < 0 ? '-' : '';
    if (abs >= 1000000) return `${sign}$${(abs / 1000000).toLocaleString('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}M`;
    if (abs >= 1000) return `${sign}$${(abs / 1000).toLocaleString('es-CO', { maximumFractionDigits: 0 })}K`;
    return `${sign}$${abs.toLocaleString('es-CO', { maximumFractionDigits: 0 })}`;
  }

  function formatPercentRatio(value, base) {
    if (!base) return value ? 'sin ingreso' : '0,0%';
    return `${(value / base * 100).toLocaleString('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
  }

  // ── Inicialización de Datos y Totales ──────────────────────────────
  function getProcessedBrandData(brand) {
    let totalIncomeCOP = 0;
    let totalIncomeUSD = 0;
    (brand.incomes || []).forEach(inc => {
      if (inc.status === 'Pendiente') return; // Excluir pendientes
      totalIncomeCOP += inc.cop || 0;
      totalIncomeUSD += inc.usd || 0;
    });

    let totalOutflowCOP = 0;
    let eventsCOP = 0;
    let incentivesCOP = 0;
    let rioCOP = 0;
    let otherCOP = 0;

    (brand.outflows || []).forEach(out => {
      const cop = out.cop || 0;
      totalOutflowCOP += cop;
      const type = (out.type || '').toLowerCase();
      if (type.includes('evento')) eventsCOP += cop;
      else if (type.includes('incentivo')) incentivesCOP += cop;
      else if (type.includes('rio')) rioCOP += cop;
      else otherCOP += cop;
    });

    const balanceCOP = totalIncomeCOP - totalOutflowCOP;
    const pctExecuted = totalIncomeCOP > 0 ? (totalOutflowCOP / totalIncomeCOP) : (totalOutflowCOP > 0 ? 1 : 0);
    const commissionCOP = balanceCOP > 0 ? balanceCOP * 0.10 : 0;

    return {
      ...brand,
      totalIncomeCOP,
      totalIncomeUSD,
      totalOutflowCOP,
      eventsCOP,
      incentivesCOP,
      rioCOP,
      otherCOP,
      balanceCOP,
      pctExecuted,
      commissionCOP
    };
  }

  function getGlobalSummary(brandsData) {
    let totalIncome = 0;
    let totalOutflow = 0;
    let totalEvents = 0;
    let totalIncentives = 0;
    let totalRio = 0;
    let totalOther = 0;

    brandsData.filter(b => !b.excludeFromConsolidated).forEach(b => {
      totalIncome += b.totalIncomeCOP;
      totalOutflow += b.totalOutflowCOP;
      totalEvents += b.eventsCOP;
      totalIncentives += b.incentivesCOP;
      totalRio += b.rioCOP;
      totalOther += b.otherCOP;
    });

    const totalBalance = totalIncome - totalOutflow;
    const pctGlobalExecuted = totalIncome > 0 ? (totalOutflow / totalIncome) : 0;
    const totalCommission = totalBalance > 0 ? totalBalance * 0.10 : 0;

    return {
      totalIncome,
      totalOutflow,
      totalBalance,
      pctGlobalExecuted,
      totalCommission,
      totalEvents,
      totalIncentives,
      totalRio,
      totalOther
    };
  }

  // ── Renderizado del HTML Principal ──────────────────────────────────
  function renderLayout(container) {
    if (!container) return;

    container.innerHTML = `
      <div class="fondos-app-wrapper">
        <!-- Header del Módulo -->
        <section class="fondos-header" aria-labelledby="fondos-page-title">
          <div class="fondos-header-main">
            <div class="fondos-title-badge">
              <span class="fondos-badge-icon">💰</span>
              <div>
                <h1 class="fondos-title" id="fondos-page-title">Fondos de Mercadeo (MDF)</h1>
                <p class="fondos-subtitle">Control presupuestal, ejecuciones por fabricante y cálculo de comisión de mercadeo</p>
              </div>
            </div>
            <div class="fondos-actions">
              <button class="fondos-btn fondos-btn-secondary" id="fondos-export-btn" type="button">
                <span>📥 Exportar Reporte Excel</span>
              </button>
            </div>
          </div>

          <!-- Banner Informativo / TRM -->
          <div class="fondos-info-bar">
            <div class="fondos-info-item">
              <span class="info-label">Origen de Datos:</span>
              <span class="info-val" id="fondos-status-val">✅ Informe final · ${escapeMarkup(state.data && state.data.sourceFileName || BASELINE_DATA.sourceFileName)}</span>
            </div>
            <div class="fondos-info-item">
              <span class="info-label">TRM Referencia:</span>
              <span class="info-val" id="fondos-trm-val">$3.500 COP/USD</span>
            </div>
            <div class="fondos-info-item">
              <span class="info-label">Acceso Restringido:</span>
              <span class="info-val status-secured">🔒 Nini · C.Estratégica · Juan Novoa · Especialista</span>
            </div>
          </div>
        </section>

        <!-- Tarjetas KPI Principales -->
        <section class="fondos-kpi-grid" id="fondos-kpis-container"></section>

        <!-- Filtros e Interacción por Marca -->
        <section class="fondos-controls-bar">
          <div class="fondos-chip-group" id="fondos-brand-chips"></div>
          <div class="fondos-category-select-wrap">
            <label for="fondos-cat-filter">Categoría de Salida:</label>
            <select id="fondos-cat-filter" class="fondos-select">
              <option value="all">Todas las Salidas</option>
              <option value="Evento">Eventos</option>
              <option value="Incentivo">Incentivos</option>
              <option value="Rio">Convención Río</option>
              <option value="Otro">Otros</option>
            </select>
          </div>
        </section>

        <!-- Sección de Gráficos de Alto Impacto -->
        <section class="fondos-charts-grid">
          <div class="fondos-card chart-card">
            <div class="fondos-card-header">
              <h3 class="fondos-card-title">📊 Presupuesto vs Ejecución por Marca</h3>
              <span class="fondos-card-tag">Comparativo COP</span>
            </div>
            <div class="chart-container">
              <canvas id="fondos-bar-chart"></canvas>
            </div>
          </div>
          <div class="fondos-card chart-card">
            <div class="fondos-card-header">
              <h3 class="fondos-card-title">🍩 Distribución de Salidas por Categoría</h3>
              <span class="fondos-card-tag">Consolidado de Gastos</span>
            </div>
            <div class="chart-container">
              <canvas id="fondos-doughnut-chart"></canvas>
            </div>
          </div>
        </section>

        <!-- Tabla Consolidada General -->
        <section class="fondos-card table-card">
          <div class="fondos-card-header">
            <div>
              <h3 class="fondos-card-title">📋 Consolidado de Fondos por Fabricante</h3>
              <p class="fondos-card-sub">Resumen de ingresos, salidas registradas, saldos disponibles y porcentaje ejecutado</p>
            </div>
          </div>
          <div class="table-responsive">
            <table class="fondos-table" id="fondos-summary-table">
              <thead>
                <tr>
                  <th>Fabricante</th>
                  <th class="num">Ingreso Total (COP)</th>
                  <th class="num">Ejecutado (COP)</th>
                  <th class="num">Saldo Disponible (COP)</th>
                  <th class="center">% Ejecutado</th>
                  <th class="num">Comisión Dir. (10%)</th>
                  <th class="center">Estado</th>
                </tr>
              </thead>
              <tbody id="fondos-summary-tbody"></tbody>
            </table>
          </div>
        </section>

        <!-- Detalle Desplegable por Marca -->
        <section class="fondos-brand-detail-section" id="fondos-brand-detail-container"></section>
      </div>
    `;
  }

  // ── Renderizado de KPIs ─────────────────────────────────────────────
  function renderKPIs(summary) {
    const container = document.getElementById('fondos-kpis-container');
    if (!container) return;

    const pctText = (summary.pctGlobalExecuted * 100).toFixed(1) + '%';

    container.innerHTML = `
      <div class="kpi-card kpi-income">
        <div class="kpi-icon-wrap">💵</div>
        <div class="kpi-body">
          <span class="kpi-label">Ingresos Totales</span>
          <div class="kpi-value">${formatCOP(summary.totalIncome)}</div>
          <span class="kpi-sub">Presupuesto confirmado por marcas</span>
        </div>
      </div>

      <div class="kpi-card kpi-outflow">
        <div class="kpi-icon-wrap">📉</div>
        <div class="kpi-body">
          <span class="kpi-label">Ejecutado (Salidas)</span>
          <div class="kpi-value">${formatCOP(summary.totalOutflow)}</div>
          <div class="kpi-progress-bar">
            <div class="kpi-progress-fill" style="width: ${Math.min(100, summary.pctGlobalExecuted * 100)}%"></div>
          </div>
          <span class="kpi-sub">${pctText} del presupuesto total</span>
        </div>
      </div>

      <div class="kpi-card kpi-balance">
        <div class="kpi-icon-wrap">💎</div>
        <div class="kpi-body">
          <span class="kpi-label">Saldo Disponible Neto</span>
          <div class="kpi-value ${summary.totalBalance < 0 ? 'text-danger' : 'text-success'}">${formatCOP(summary.totalBalance)}</div>
          <span class="kpi-sub">Fondo disponible para reasignar</span>
        </div>
      </div>

      <div class="kpi-card kpi-commission">
        <div class="kpi-icon-wrap">🎯</div>
        <div class="kpi-body">
          <span class="kpi-label">Comisión Dir. Mercadeo</span>
          <div class="kpi-value text-accent">${formatCOP(summary.totalCommission)}</div>
          <span class="kpi-sub">10% sobre el saldo neto consolidado</span>
        </div>
      </div>
    `;
  }

  // ── Renderizado de Chips de Marca ────────────────────────────────────
  function renderBrandChips(brands) {
    const container = document.getElementById('fondos-brand-chips');
    if (!container) return;

    let html = `
      <button class="brand-chip ${state.activeBrandId === 'all' ? 'active' : ''}" data-id="all">
        <span>🌐 Todas las Marcas</span>
      </button>
    `;

    brands.forEach(b => {
      const activeClass = state.activeBrandId === b.id ? 'active' : '';
      html += `
        <button class="brand-chip ${activeClass}" data-id="${b.id}" style="--chip-color: ${b.color}">
          <span class="chip-dot" style="background: ${b.color}"></span>
          <span>${b.name}</span>
        </button>
      `;
    });

    container.innerHTML = html;

    container.querySelectorAll('.brand-chip').forEach(btn => {
      btn.addEventListener('click', () => {
        state.activeBrandId = btn.getAttribute('data-id');
        updateDashboardView();
      });
    });
  }

  // ── Renderizado de Tabla Consolidada ────────────────────────────────
  function renderSummaryTable(processedBrands, summary) {
    const tbody = document.getElementById('fondos-summary-tbody');
    if (!tbody) return;

    let html = '';

    processedBrands.forEach(b => {
      const isSelected = state.activeBrandId === b.id;
      const pctVal = (b.pctExecuted * 100).toFixed(1);

      let statusBadge = '<span class="status-pill pill-normal">En Ejecución</span>';
      if (b.excludeFromConsolidated) statusBadge = '<span class="status-pill pill-pending">Control separado</span>';
      else if (b.balanceCOP < 0) statusBadge = '<span class="status-pill pill-danger">Déficit</span>';
      else if (b.pctExecuted >= 0.99) statusBadge = '<span class="status-pill pill-complete">100% Ejecutado</span>';
      else if (b.pctExecuted === 0) statusBadge = '<span class="status-pill pill-pending">Sin Ejecutar</span>';

      html += `
        <tr class="${isSelected ? 'selected-row' : ''}" data-brand-id="${b.id}">
          <td class="brand-name-cell">
            <span class="brand-badge-dot" style="background:${b.color}"></span>
            <strong>${b.name}</strong>
          </td>
          <td class="num">${formatCOP(b.totalIncomeCOP)}</td>
          <td class="num">${formatCOP(b.totalOutflowCOP)}</td>
          <td class="num ${b.balanceCOP < 0 ? 'text-danger font-bold' : ''}">${formatCOP(b.balanceCOP)}</td>
          <td class="center">
            <div class="table-progress-wrap">
              <span class="table-progress-text">${pctVal}%</span>
              <div class="table-progress-bar">
                <div class="table-progress-fill" style="width:${Math.min(100, Math.max(0, pctVal))}%; background:${b.color}"></div>
              </div>
            </div>
          </td>
          <td class="num text-accent font-bold">${formatCOP(b.commissionCOP)}</td>
          <td class="center">${statusBadge}</td>
        </tr>
      `;
    });

    // Fila Total
    const pctTotalText = (summary.pctGlobalExecuted * 100).toFixed(1) + '%';
    html += `
      <tr class="total-row">
        <td><strong>TOTAL CONSOLIDADO</strong></td>
        <td class="num"><strong>${formatCOP(summary.totalIncome)}</strong></td>
        <td class="num"><strong>${formatCOP(summary.totalOutflow)}</strong></td>
        <td class="num ${summary.totalBalance < 0 ? 'text-danger' : 'text-success'}"><strong>${formatCOP(summary.totalBalance)}</strong></td>
        <td class="center">
          <strong>${pctTotalText}</strong>
        </td>
        <td class="num text-accent font-bold"><strong>${formatCOP(summary.totalCommission)}</strong></td>
        <td class="center"><span class="status-pill pill-gold">Consolidado</span></td>
      </tr>
    `;

    tbody.innerHTML = html;

    tbody.querySelectorAll('tr[data-brand-id]').forEach(row => {
      row.addEventListener('click', () => {
        const id = row.getAttribute('data-brand-id');
        state.activeBrandId = state.activeBrandId === id ? 'all' : id;
        updateDashboardView();
      });
    });
  }

  // ── Renderizado de Gráficos (Chart.js) ──────────────────────────────
  function renderCharts(processedBrands, summary) {
    if (typeof Chart === 'undefined') {
      console.warn('[FONDOS] Chart.js no cargado');
      return;
    }

    // 1. Gráfico de Barras: Presupuesto vs Salida vs Saldo
    const ctxBar = document.getElementById('fondos-bar-chart');
    if (ctxBar) {
      if (state.charts.barChart) state.charts.barChart.destroy();

      const labels = processedBrands.map(b => b.name);
      const incomes = processedBrands.map(b => b.totalIncomeCOP);
      const outflows = processedBrands.map(b => b.totalOutflowCOP);
      const balances = processedBrands.map(b => b.balanceCOP);
      const barValueLabels = {
        id: 'fondosBarValueLabels',
        afterDatasetsDraw(chart) {
          const chartCtx = chart.ctx;
          const textColor = getComputedStyle(document.documentElement).getPropertyValue('--text').trim() || '#172033';
          chartCtx.save();
          chartCtx.textAlign = 'center';
          chartCtx.textBaseline = 'middle';
          chartCtx.lineJoin = 'round';
          chartCtx.strokeStyle = getComputedStyle(document.documentElement).getPropertyValue('--card').trim() || 'rgba(255,255,255,.92)';
          chartCtx.lineWidth = 3;
          chart.data.datasets.forEach((dataset, datasetIndex) => {
            const meta = chart.getDatasetMeta(datasetIndex);
            meta.data.forEach((element, index) => {
              const value = Number(dataset.data[index] || 0);
              if (!value) return;
              const income = Number(incomes[index] || 0);
              const position = element.tooltipPosition();
              const isNegative = value < 0;
              const valueY = position.y + (isNegative ? 13 : -13);
              chartCtx.font = '700 9px Plus Jakarta Sans, sans-serif';
              chartCtx.strokeText(formatCompactCOP(value), position.x, valueY);
              chartCtx.fillStyle = textColor;
              chartCtx.fillText(formatCompactCOP(value), position.x, valueY);
              if (datasetIndex > 0) {
                const percentY = valueY + (isNegative ? 11 : -11);
                const percentText = formatPercentRatio(value, income);
                chartCtx.font = '800 8px Plus Jakarta Sans, sans-serif';
                chartCtx.strokeText(percentText, position.x, percentY);
                chartCtx.fillStyle = datasetIndex === 1 ? '#DC2626' : '#059669';
                chartCtx.fillText(percentText, position.x, percentY);
              }
            });
          });
          chartCtx.restore();
        }
      };

      state.charts.barChart = new Chart(ctxBar, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: 'Ingreso (COP)',
              data: incomes,
              backgroundColor: 'rgba(59, 130, 246, 0.75)',
              borderColor: '#3B82F6',
              borderWidth: 1,
              borderRadius: 4
            },
            {
              label: 'Ejecutado (COP)',
              data: outflows,
              backgroundColor: 'rgba(239, 68, 68, 0.75)',
              borderColor: '#EF4444',
              borderWidth: 1,
              borderRadius: 4
            },
            {
              label: 'Saldo (COP)',
              data: balances,
              backgroundColor: 'rgba(16, 185, 129, 0.75)',
              borderColor: '#10B981',
              borderWidth: 1,
              borderRadius: 4
            }
          ]
        },
        plugins: [barValueLabels],
        options: {
          responsive: true,
          maintainAspectRatio: false,
          layout: { padding: { top: 34, bottom: 26, left: 4, right: 4 } },
          plugins: {
            legend: { position: 'top', labels: { color: 'var(--text, #e2e8f0)', font: { family: 'Plus Jakarta Sans' } } },
            tooltip: {
              callbacks: {
                label: function (ctx) {
                  const income = incomes[ctx.dataIndex] || 0;
                  const percent = ctx.datasetIndex === 0 ? '100,0%' : formatPercentRatio(ctx.raw, income);
                  return ` ${ctx.dataset.label}: ${formatCOP(ctx.raw)} · ${percent}`;
                }
              }
            }
          },
          scales: {
            x: { ticks: { color: 'var(--text2, #94a3b8)' }, grid: { display: false } },
            y: {
              ticks: {
                color: 'var(--text2, #94a3b8)',
                callback: function (val) {
                  return '$' + (val / 1e6).toFixed(0) + 'M';
                }
              },
              grid: { color: 'rgba(255, 255, 255, 0.05)' }
            }
          }
        }
      });
    }

    // 2. Gráfico de Dona: Distribución de Salidas
    const ctxDoughnut = document.getElementById('fondos-doughnut-chart');
    if (ctxDoughnut) {
      if (state.charts.doughnutChart) state.charts.doughnutChart.destroy();

      const championControlTotal = processedBrands
        .filter(brand => brand.excludeFromConsolidated)
        .reduce((total, brand) => total + brand.totalOutflowCOP, 0);
      const doughnutLabels = ['Eventos', 'Incentivos', 'Convención Río', 'Champion'];
      const doughnutColors = ['#8B5CF6', '#F59E0B', '#06B6D4', '#64748B'];
      const doughnutValues = [
        summary.totalEvents,
        summary.totalIncentives,
        summary.totalRio,
        summary.totalOther + championControlTotal
      ];
      const doughnutTotal = doughnutValues.reduce((total, value) => total + value, 0);
      const doughnutValueLabels = {
        id: 'fondosDoughnutValueLabels',
        afterDatasetsDraw(chart) {
          const chartCtx = chart.ctx;
          const textColor = getComputedStyle(document.documentElement).getPropertyValue('--text').trim() || '#172033';
          chartCtx.save();
          chartCtx.textAlign = 'center';
          chartCtx.textBaseline = 'middle';
          chart.getDatasetMeta(0).data.forEach((arc, index) => {
            const value = Number(doughnutValues[index] || 0);
            const percent = doughnutTotal ? value / doughnutTotal * 100 : 0;
            if (percent < 5) return;
            const position = arc.tooltipPosition();
            chartCtx.strokeStyle = 'rgba(255,255,255,.92)';
            chartCtx.lineWidth = 3;
            chartCtx.font = '800 9px Plus Jakarta Sans, sans-serif';
            chartCtx.strokeText(`${percent.toLocaleString('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`, position.x, position.y - 6);
            chartCtx.fillStyle = '#0F172A';
            chartCtx.fillText(`${percent.toLocaleString('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`, position.x, position.y - 6);
            chartCtx.font = '700 8px Plus Jakarta Sans, sans-serif';
            chartCtx.strokeText(formatCompactCOP(value), position.x, position.y + 6);
            chartCtx.fillText(formatCompactCOP(value), position.x, position.y + 6);
          });
          const { left, right, top, bottom } = chart.chartArea;
          chartCtx.fillStyle = textColor;
          chartCtx.font = '700 10px Plus Jakarta Sans, sans-serif';
          chartCtx.fillText('Salidas visibles', (left + right) / 2, (top + bottom) / 2 - 12);
          chartCtx.font = '800 14px Plus Jakarta Sans, sans-serif';
          chartCtx.fillText(formatCompactCOP(doughnutTotal), (left + right) / 2, (top + bottom) / 2 + 5);
          chartCtx.font = '600 8px Plus Jakarta Sans, sans-serif';
          chartCtx.fillText('incluye Champion', (left + right) / 2, (top + bottom) / 2 + 20);
          chartCtx.restore();
        }
      };

      state.charts.doughnutChart = new Chart(ctxDoughnut, {
        type: 'doughnut',
        data: {
          labels: doughnutLabels,
          datasets: [{
            data: doughnutValues,
            backgroundColor: doughnutColors,
            borderWidth: 2,
            borderColor: 'var(--bg-card, #1e293b)'
          }]
        },
        plugins: [doughnutValueLabels],
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'right',
              labels: {
                color: 'var(--text, #e2e8f0)',
                font: { family: 'Plus Jakarta Sans', size: 10 },
                generateLabels(chart) {
                  return doughnutLabels.map((label, index) => {
                    const value = doughnutValues[index] || 0;
                    const percent = doughnutTotal ? value / doughnutTotal * 100 : 0;
                    return {
                      text: `${label} · ${formatCompactCOP(value)} · ${percent.toLocaleString('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`,
                      fillStyle: doughnutColors[index],
                      strokeStyle: doughnutColors[index],
                      lineWidth: 1,
                      hidden: !chart.getDataVisibility(index),
                      index
                    };
                  });
                }
              }
            },
            tooltip: {
              callbacks: {
                label: function (ctx) {
                  const percent = doughnutTotal ? ctx.raw / doughnutTotal * 100 : 0;
                  return ` ${ctx.label}: ${formatCOP(ctx.raw)} · ${percent.toLocaleString('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
                }
              }
            }
          },
          cutout: '65%'
        }
      });
    }
  }

  // ── Renderizado de Detalle por Marca ────────────────────────────────
  function renderBrandDetail(processedBrands) {
    const container = document.getElementById('fondos-brand-detail-container');
    if (!container) return;

    const filtered = state.activeBrandId === 'all'
      ? processedBrands
      : processedBrands.filter(b => b.id === state.activeBrandId);

    let html = '';

    filtered.forEach(b => {
      const filteredOutflows = (b.outflows || []).filter(out => {
        if (state.activeCategory === 'all') return true;
        return (out.type || '').toLowerCase().includes(state.activeCategory.toLowerCase());
      });

      const hasNotes = Boolean(b.notes);

      html += `
        <article class="fondos-brand-card" style="--brand-color:${b.color}">
          <header class="brand-card-header" style="background:${b.bgAlpha}">
            <div class="brand-header-title">
              <span class="brand-badge-pill" style="background:${b.color}">${b.name}</span>
              <span class="brand-subtitle font-mono">TRM: $${b.incomes[0]?.trm ? Number(b.incomes[0].trm).toLocaleString('es-CO') : '3.500'} COP</span>
            </div>
            <div class="brand-header-stats">
              <div class="stat-box">
                <span class="lbl">Ingresos</span>
                <span class="val">${formatCOP(b.totalIncomeCOP)}</span>
              </div>
              <div class="stat-box">
                <span class="lbl">Ejecutado</span>
                <span class="val">${formatCOP(b.totalOutflowCOP)}</span>
              </div>
              <div class="stat-box">
                <span class="lbl">Saldo</span>
                <span class="val ${b.balanceCOP < 0 ? 'text-danger' : 'text-success'}">${formatCOP(b.balanceCOP)}</span>
              </div>
            </div>
          </header>

          ${hasNotes ? `
            <div class="brand-notes-callout">
              <span class="notes-icon">📌</span>
              <p class="notes-text">${b.notes}</p>
            </div>
          ` : ''}

          <div class="brand-sections-grid">
            <!-- 1. Ingreso de Fondos -->
            <div class="brand-subcard">
              <h4 class="subcard-title">1. Ingreso de Fondos</h4>
              ${b.incomes && b.incomes.length ? `
                <table class="fondos-subtable">
                  <thead>
                    <tr>
                      <th>Período / Origen</th>
                      <th class="num">USD</th>
                      <th class="num">TRM</th>
                      <th class="num">COP</th>
                      <th>Observación</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${b.incomes.map(inc => `
                      <tr class="${inc.status === 'Pendiente' ? 'row-pending' : ''}">
                        <td><strong>${inc.period || inc.origen || '-'}</strong></td>
                        <td class="num font-mono">${formatUSD(inc.usd)}</td>
                        <td class="num font-mono">$${inc.trm ? Number(inc.trm).toLocaleString('es-CO') : '-'}</td>
                        <td class="num font-mono">${formatCOP(inc.cop)}</td>
                        <td><span class="obs-tag">${inc.obs || inc.status || '-'}</span></td>
                      </tr>
                    `).join('')}
                  </tbody>
                </table>
              ` : `
                <div class="empty-subcard">Sin ingreso de fondos registrado</div>
              `}
            </div>

            <!-- 2. Salidas (Ejecución) -->
            <div class="brand-subcard">
              <h4 class="subcard-title">2. Salidas (Ejecución)</h4>
              ${filteredOutflows.length ? `
                <table class="fondos-subtable">
                  <thead>
                    <tr>
                      <th>Tipo</th>
                      <th>Concepto</th>
                      <th class="center">Cupos / Fecha</th>
                      <th class="num">COP</th>
                      <th>Observación</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${filteredOutflows.map(out => `
                      <tr>
                        <td><span class="type-pill type-${(out.type||'otro').toLowerCase()}">${out.type || 'Otro'}</span></td>
                        <td><strong>${out.concept || '-'}</strong></td>
                        <td class="center">${out.cupos ? `${out.cupos} cupos` : (out.obs && out.obs.includes('2026') ? out.obs : '-')}</td>
                        <td class="num font-mono text-outflow">${formatCOP(out.cop)}</td>
                        <td><span class="obs-tag">${out.obs || '-'}</span></td>
                      </tr>
                    `).join('')}
                  </tbody>
                </table>
              ` : `
                <div class="empty-subcard">No hay salidas registradas para el filtro seleccionado</div>
              `}
            </div>
          </div>

          <!-- Resumen de Marca -->
          <footer class="brand-card-footer">
            <div class="brand-footer-summary">
              <span><strong>Saldo Disponible:</strong> ${formatCOP(b.balanceCOP)}</span>
              <span><strong>Comisión Dir. Mercadeo (10%):</strong> <strong class="text-accent">${formatCOP(b.commissionCOP)}</strong></span>
            </div>
          </footer>
        </article>
      `;
    });

    container.innerHTML = html;
  }

  // ── Actualizador General de la Vista ─────────────────────────────
  function updateDashboardView() {
    if (!state.data) return;

    const processedBrands = state.data.brands.map(getProcessedBrandData);
    const summary = getGlobalSummary(processedBrands);

    renderKPIs(summary);
    renderBrandChips(processedBrands);
    renderSummaryTable(processedBrands, summary);
    renderCharts(processedBrands, summary);
    renderBrandDetail(processedBrands);
  }

  // ── Carga y Procesamiento de Excel ────────────────────────────────
  function handleExcelUpload(file) {
    if (!file || typeof XLSX === 'undefined') return;

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        console.log('[FONDOS] Excel leído con éxito. Hojas:', workbook.SheetNames);

        document.getElementById('fondos-status-val').textContent = `Cargado localmente: ${file.name}`;
        alert(`¡Archivo "${file.name}" cargado exitosamente! El dashboard se actualizará en memoria.`);
      } catch (err) {
        console.error('[FONDOS] Error procesando Excel:', err);
        alert('Ocurrió un error al leer el archivo Excel.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ── Exportación a Excel ─────────────────────────────────────────────
  function exportToExcel() {
    if (typeof XLSX === 'undefined' || !state.data) return;

    const processedBrands = state.data.brands.map(getProcessedBrandData);
    const summary = getGlobalSummary(processedBrands);

    const summaryRows = processedBrands.map(b => ({
      'Fabricante': b.name,
      'Ingreso Total (COP)': b.totalIncomeCOP,
      'Ejecutado (COP)': b.totalOutflowCOP,
      'Saldo Disponible (COP)': b.balanceCOP,
      '% Ejecutado': (b.pctExecuted * 100).toFixed(1) + '%',
      'Comisión Dir. Mercadeo 10% (COP)': b.commissionCOP,
      'Observaciones': b.notes || ''
    }));

    summaryRows.push({
      'Fabricante': 'TOTAL CONSOLIDADO',
      'Ingreso Total (COP)': summary.totalIncome,
      'Ejecutado (COP)': summary.totalOutflow,
      'Saldo Disponible (COP)': summary.totalBalance,
      '% Ejecutado': (summary.pctGlobalExecuted * 100).toFixed(1) + '%',
      'Comisión Dir. Mercadeo 10% (COP)': summary.totalCommission,
      'Observaciones': 'Consolidado General'
    });

    const wb = XLSX.utils.book_new();
    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Consolidado');

    XLSX.writeFile(wb, `Fondos_Mercadeo_Forecast_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  // ── Inicialización del Módulo ───────────────────────────────────────
  function init() {
    const container = document.getElementById('page-fondos');
    if (!container) return;

    if (!canAccess()) {
      container.innerHTML = `
        <div class="fondos-access-denied">
          <div class="denied-card">
            <span class="denied-icon">🔒</span>
            <h2>Acceso Restringido</h2>
            <p>La vista <strong>Fondos Marketing</strong> está habilitada exclusivamente para Gerencia.</p>
          </div>
        </div>
      `;
      return;
    }

    // Reconstruir siempre: otras actualizaciones de la aplicación pueden vaciar
    // el contenedor mientras el módulo permanece cargado en memoria.
    Object.values(state.charts).forEach(chart => {
      if(chart && typeof chart.destroy === 'function') chart.destroy();
    });
    state.charts.barChart = null;
    state.charts.doughnutChart = null;
    if (!state.data) state.data = JSON.parse(JSON.stringify(BASELINE_DATA));
    renderLayout(container);

    const exportBtn = document.getElementById('fondos-export-btn');
    if (exportBtn) exportBtn.addEventListener('click', exportToExcel);

    const catSelect = document.getElementById('fondos-cat-filter');
    if (catSelect) {
      catSelect.addEventListener('change', (e) => {
        state.activeCategory = e.target.value;
        updateDashboardView();
      });
    }

    updateDashboardView();
    if (!container.querySelector('.fondos-app-wrapper')) {
      throw new Error('El tablero de Fondos Marketing no quedó insertado en la página.');
    }
  }

  window.FondosMarketingModule = {
    init,
    canAccess,
    loadFromSharePoint,
    updateDashboardView
  };

})();
