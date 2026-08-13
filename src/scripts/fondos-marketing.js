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

  // Datos Base Consolidados (Extraídos y validados del libro FONDOS DE MERCADEO NINI.xlsx)
  const BASELINE_DATA = {
    trmDefault: 3500,
    lastUpdate: '2026-08-13',
    brands: [
      {
        id: 'hp',
        name: 'HP',
        shortName: 'Hp',
        color: '#0096D6',
        bgAlpha: 'rgba(0, 150, 214, 0.18)',
        incomes: [
          { period: '26Q1', usd: 2200, trm: 3500, cop: 7700000, obs: 'Fondo asignado Q1' },
          { period: '26Q1', usd: 1000, trm: 3500, cop: 3500000, obs: 'Adicional Q1' },
          { period: '26Q1', usd: 1500, trm: 3500, cop: 5250000, obs: 'Incentivos Q1' },
          { period: '26Q1', usd: 1500, trm: 3500, cop: 5250000, obs: 'Eventos Q1' },
          { period: '26Q2', usd: 3000, trm: 3500, cop: 10500000, obs: 'Fondo base Q2' },
          { period: '26Q2', usd: 3000, trm: 3500, cop: 10500000, obs: 'Suministros Q2' },
          { period: '26Q2', usd: 1000, trm: 3500, cop: 3500000, obs: 'Adicional Q2' },
          { period: '26Q3', usd: 3000, trm: 3500, cop: 10500000, obs: 'Fondo base Q3' },
          { period: '26Q3', usd: 1000, trm: 3500, cop: 3500000, obs: 'Adicional Q3' }
        ],
        outflows: [
          { type: 'Evento', concept: 'HP Suministros - Evento', cupos: null, cop: 2710400, obs: 'Evento presencial suministros' },
          { type: 'Evento', concept: 'Liga Z - Capacitación', cupos: null, cop: 486000, obs: 'Capacitación fuerza de ventas' },
          { type: 'Evento', concept: 'HP Workstation - Hotel Hilton', cupos: null, cop: 1322585, obs: 'Lanzamiento Workstation' },
          { type: 'Evento', concept: 'HP Nueva era de la tecnología - Hotel', cupos: null, cop: 5943183, obs: 'Convención hotelera' },
          { type: 'Incentivo', concept: 'Concurso Supplies', cupos: null, cop: 2500000, obs: 'Premios canal suministros' },
          { type: 'Incentivo', concept: 'Concurso venta portafolio HP', cupos: null, cop: 3000000, obs: 'Incentivo vendedores' },
          { type: 'Incentivo', concept: 'Campaña Suministros', cupos: null, cop: 6000000, obs: 'Campaña comercial' },
          { type: 'Incentivo', concept: 'Camisetas (obsequio gerencia comercial y direcciones)', cupos: null, cop: 1500000, obs: 'Dotación corporativa' },
          { type: 'Incentivo', concept: 'Incentivo Suministros (Bono Adidas ganadora Dilma)', cupos: null, cop: 300000, obs: 'Bono regalo' },
          { type: 'Incentivo', concept: 'Nintendo', cupos: null, cop: 2600000, obs: 'Premio meta de ventas' },
          { type: 'Rio', concept: 'Rio (20 al 24 mayo) - cupos @ $6.500.000', cupos: 3, cop: 19500000, obs: '3 cupos convención Río' }
        ],
        notes: 'TRM estándar $3.500. Los montos y conceptos de ejecuciones se encuentran 100% alineados y auditados.'
      },
      {
        id: 'dell',
        name: 'DELL',
        shortName: 'Dell',
        color: '#0076CE',
        bgAlpha: 'rgba(0, 118, 206, 0.18)',
        incomes: [
          { period: '2026 Q3', usd: 1414.49, trm: 3500, cop: 4950715, obs: 'Agosto - Octubre 2026' },
          { period: '2026 Q4', usd: 967, trm: 3500, cop: 3384500, obs: 'Noviembre 1 - Enero 30 (cierra ene 2027)' },
          { period: '2027 Q1', usd: 2864, trm: 3500, cop: 10024000, obs: 'Enero 31 - Mayo 1' },
          { period: '2027 Q2', usd: 1811, trm: 3500, cop: 6338500, obs: 'Mayo 2 - Julio 31' },
          { period: 'Proposal', usd: 9500, trm: 3500, cop: 33250000, obs: 'Fondos Proposal - Generación de demanda (factura Dell)' }
        ],
        outflows: [
          { type: 'Evento', concept: 'Infraestructura', cupos: null, cop: 8874800, obs: '24 de abril 2026' },
          { type: 'Evento', concept: 'Tecnología cómputo', cupos: null, cop: 6452650, obs: '17 de julio 2026' },
          { type: 'Incentivo', concept: 'Televisor, Apple TV, bonos H&M', cupos: null, cop: 3000000, obs: 'Premios fidelización' },
          { type: 'Incentivo', concept: 'Play Station', cupos: null, cop: 2600000, obs: 'Premio mejor ejecutivo' },
          { type: 'Rio', concept: 'Rio (20 al 24 mayo) - cupos @ $6.500.000', cupos: 2, cop: 13000000, obs: '2 cupos convención Río' }
        ],
        notes: "⚠️ ATENCIÓN: El ingreso incluye 'Fondos Proposal' (USD 9.500 = $33.250.000 COP) aún no ejecutado. Si se excluye este rubro, el saldo real disponible de Dell pasa a ser -$9.229.735 COP."
      },
      {
        id: 'lenovo',
        name: 'LENOVO',
        shortName: 'Lenovo',
        color: '#E2231A',
        bgAlpha: 'rgba(226, 35, 26, 0.18)',
        incomes: [
          { period: 'Q2', usd: 1906, trm: 3500, cop: 6671000, status: 'Confirmado', obs: 'Facturado' },
          { period: 'Q3', usd: 2997, trm: 3500, cop: 10489500, status: 'Confirmado', obs: 'Facturado' },
          { period: 'Q4', usd: 3349, trm: 3500, cop: 11721500, status: 'Pendiente', obs: 'Sin factura Provexpress (pendiente por legalizar)' }
        ],
        outflows: [
          { type: 'Evento', concept: 'Infraestructura', cupos: null, cop: 4408950, obs: 'Evento técnico Lenovo' },
          { type: 'Incentivo', concept: 'Lenovo Legion Go S', cupos: null, cop: 2500000, obs: 'Incentivo comercial' },
          { type: 'Incentivo', concept: 'Bono María Paola Briceño', cupos: null, cop: 500000, obs: 'Reconocimiento especial' },
          { type: 'Rio', concept: 'Rio (20 al 24 mayo) - cupos @ $6.500.000', cupos: 1.5, cop: 9750000, obs: '1.5 cupos convención Río' }
        ],
        notes: '📌 NOTA: El Q4 (USD 3.349 = $11.721.500 COP) NO se contabiliza en el ingreso confirmado por falta de factura Provexpress. Al facturarse se convertirá en saldo disponible.'
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
        notes: 'El encabezado original indicaba 2 cupos de Río pero el cálculo usó 1.5 cupos ($9.750.000), logrando un cierre perfecto con saldo $0.'
      },
      {
        id: 'microsoft',
        name: 'MICROSOFT',
        shortName: 'Microsoft',
        color: '#107C41',
        bgAlpha: 'rgba(16, 124, 65, 0.18)',
        incomes: [
          { period: 'LOL', usd: 3000, trm: 3500, cop: 10500000, obs: 'Fondo mayorista LOL (TRM 3.500)' },
          { period: 'INGRAM', usd: 4876, trm: 3500, cop: 17066000, obs: 'Fondo mayorista Ingram Micro (TRM 3.500)' }
        ],
        outflows: [],
        notes: '💡 100% DISPONIBLE: No registra ejecuciones ni salidas cargadas a la fecha ($27.566.000 COP disponibles).'
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
        notes: 'Usa TRM implícita de $3.333,33 COP/USD. Si se aplicara la TRM estándar de $3.500, el ingreso registraría $10.500.000 COP (+$500.000).'
      },
      {
        id: 'champion',
        name: 'CHAMPION',
        shortName: 'Champion',
        color: '#F59E0B',
        bgAlpha: 'rgba(245, 158, 11, 0.18)',
        incomes: [],
        outflows: [
          { brand: 'HP', period: 'Q1', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion HP Q1' },
          { brand: 'HP', period: 'Q2', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion HP Q2' },
          { brand: 'DELL', period: 'Q1', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion DELL Q1' },
          { brand: 'DELL', period: 'Q2', usd: 2000, trm: 3500, cop: 7000000, type: 'Otro', concept: 'Ejecución Champion DELL Q2' }
        ],
        notes: '🚨 DÉFICIT: Registra egresos por USD 8.000 ($28.000.000 COP) sin ingreso de fondos registrado en este libro, generando saldo negativo (-$28.000.000 COP).'
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

  // ── Permisos de Acceso ──────────────────────────────────────────────
  function canAccess() {
    if (!window.CURRENT_USER || !CURRENT_USER.role) return true; // Vista local / dev
    const role = CURRENT_USER.role;
    if (role === 'gerencia' || role === 'gerencia_director') return true;

    const userEmail = String(CURRENT_USER.email || '').toLowerCase().trim();
    const candidates = (CURRENT_USER.candidates || []).map(c => String(c || '').toLowerCase().trim());
    const allUserEmails = [userEmail, ...candidates];

    return ALLOWED_USERS.some(allowed => {
      const allowedClean = allowed.toLowerCase().trim();
      const prefix = allowedClean.split('@')[0];
      return allUserEmails.some(e => e === allowedClean || e.startsWith(prefix));
    });
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

    brandsData.forEach(b => {
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
        <header class="fondos-header">
          <div class="fondos-header-main">
            <div class="fondos-title-badge">
              <span class="fondos-badge-icon">💰</span>
              <div>
                <h1 class="fondos-title">Fondos de Mercadeo (MDF)</h1>
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
              <span class="info-val" id="fondos-status-val">✅ Carga Automática · FONDOS DE MERCADEO NINI.xlsx</span>
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
        </header>

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
      if (b.balanceCOP < 0) statusBadge = '<span class="status-pill pill-danger">Déficit</span>';
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
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top', labels: { color: 'var(--text, #e2e8f0)', font: { family: 'Plus Jakarta Sans' } } },
            tooltip: {
              callbacks: {
                label: function (ctx) {
                  return ` ${ctx.dataset.label}: ${formatCOP(ctx.raw)}`;
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

      state.charts.doughnutChart = new Chart(ctxDoughnut, {
        type: 'doughnut',
        data: {
          labels: ['Eventos', 'Incentivos', 'Convención Río', 'Otros'],
          datasets: [{
            data: [summary.totalEvents, summary.totalIncentives, summary.totalRio, summary.totalOther],
            backgroundColor: [
              '#8B5CF6',
              '#F59E0B',
              '#06B6D4',
              '#64748B'
            ],
            borderWidth: 2,
            borderColor: 'var(--bg-card, #1e293b)'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'right', labels: { color: 'var(--text, #e2e8f0)', font: { family: 'Plus Jakarta Sans' } } },
            tooltip: {
              callbacks: {
                label: function (ctx) {
                  return ` ${ctx.label}: ${formatCOP(ctx.raw)}`;
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
    state.data = JSON.parse(JSON.stringify(BASELINE_DATA));
    const container = document.getElementById('page-fondos');
    if (!container) return;

    if (!canAccess()) {
      container.innerHTML = `
        <div class="fondos-access-denied">
          <div class="denied-card">
            <span class="denied-icon">🔒</span>
            <h2>Acceso Restringido</h2>
            <p>La vista <strong>Fondos Marketing</strong> está habilitada exclusivamente para las direcciones autorizadas (Nini Beltrán, C. Estratégica, Juan Novoa y Especialista Preventa).</p>
          </div>
        </div>
      `;
      return;
    }

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
  }

  window.FondosMarketingModule = {
    init,
    canAccess,
    updateDashboardView
  };

})();
