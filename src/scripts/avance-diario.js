// ════════════════════════════════════════════════════════════════════════
//  FORECAST 2026 - MÓDULO DIARIO DE CUOTA Y AVANCE COMERCIAL
// ════════════════════════════════════════════════════════════════════════

(function () {
  const UTILIDAD_API_AUTH_URL = "http://152.200.146.226:50010/api/getKey";
  const UTILIDAD_API_DATA_URL = "http://152.200.146.226:50010/consultas/api/consultaUtilidadComercialesDashboardPBI";
  const UTILIDAD_CREDS = { username: "powerbi", password: "3xpress#2025" };

  let tokenCache = null;
  let tokenExpiration = 0;

  // Formateadores
  function formatCOP(val) {
    if (val === null || val === undefined || isNaN(val)) return "$0 COP";
    return new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(val);
  }

  function formatNumber(num) {
    if (num === null || num === undefined || isNaN(num)) return "0";
    return new Intl.NumberFormat("es-CO").format(Math.round(num));
  }

  function getMonthDateRange() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");

    return {
      fechaInicial: `${year}-${month}-01`,
      fechaFinal: `${year}-${month}-${day}`
    };
  }

  async function getApiToken() {
    if (tokenCache && Date.now() < tokenExpiration) return tokenCache;

    const res = await fetch(UTILIDAD_API_AUTH_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(UTILIDAD_CREDS)
    });

    if (!res.ok) throw new Error("Fallo la autenticación con la API de Utilidad");
    const data = await res.json();
    tokenCache = data.token || data.access_token;
    tokenExpiration = Date.now() + 55 * 60 * 1000; // 55 mins cache
    return tokenCache;
  }

  async function fetchVentasDiarias() {
    const { fechaInicial, fechaFinal } = getMonthDateRange();
    const token = await getApiToken();

    const res = await fetch(UTILIDAD_API_DATA_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      body: JSON.stringify({
        Fecha_Inicial: fechaInicial,
        Fecha_Final: fechaFinal,
        Tipo_Utilidad: "venta"
      })
    });

    if (!res.ok) throw new Error("No se lograron obtener las ventas de la API");
    const json = await res.json();
    return json.response || [];
  }

  // Renderizar la UI según el usuario conectado
  async function initAvanceDiario() {
    const currentUser = window.CURRENT_USER || (sessionStorage.getItem("forecast_user") ? JSON.parse(sessionStorage.getItem("forecast_user")) : null);
    if (!currentUser) return;

    try {
      const ventas = await fetchVentasDiarias();
      const structure = window.FORECAST_STRUCTURE || {};
      const role = currentUser.role;

      // Renderizar paneles en las páginas del Forecast
      renderExecutiveAvance(currentUser, ventas, structure);
      renderDirectorAvance(currentUser, ventas, structure);
      renderGerenciaAvance(currentUser, ventas, structure);
    } catch (e) {
      console.warn("[AVANCE DIARIO] Error al cargar:", e.message);
    }
  }

  // A. VISTA EJECUTIVO / COMERCIAL
  function renderExecutiveAvance(user, ventas, structure) {
    const container = document.getElementById("ejecutivo-avance-panel");
    if (!container) return;

    const normUserEmail = (user.email || "").toLowerCase().trim();
    const userDisplayName = (user.name || "").toLowerCase().trim();

    // Buscar coincidencia en la API de ventas
    const matchVenta = ventas.find(v => {
      const desc = String(v.Descripcion || "").toLowerCase().trim();
      return desc.includes(userDisplayName.split(" ")[0]) || userDisplayName.includes(desc.split(" ")[0]);
    });

    // Buscar cuota en la estructura
    const cuotaConfig = 18000000; // Cuota por defecto base
    const avanceUtilidad = matchVenta ? Number(matchVenta.Valor_Utilidad) || 0 : 0;
    const mercancia = matchVenta ? Number(matchVenta.Valor_Mercancia) || 0 : 0;
    const pctCumplimiento = cuotaConfig > 0 ? (avanceUtilidad / cuotaConfig) * 100 : 0;

    container.innerHTML = `
      <div style="background: var(--card, #1E293B); border: 1px solid var(--border, rgba(255,255,255,0.1)); border-radius: 12px; padding: 20px; margin-bottom: 24px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
          <div>
            <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: var(--corp-cyan, #38BDF8);">Mi Avance Diario y Cuota Mensual</div>
            <h3 style="margin: 4px 0 0; font-size: 18px; color: var(--text, #F8FAFC);">${user.name || "Ejecutivo Comercial"}</h3>
          </div>
          <span style="background: rgba(56,189,248,0.15); color: #38BDF8; border: 1px solid rgba(56,189,248,0.3); padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 700;">
            ${pctCumplimiento.toFixed(1)}% Avance
          </span>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 16px;">
          <div style="background: rgba(255,255,255,0.03); border-left: 3px solid #3B82F6; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: var(--text2, #94A3B8);">Cuota Mensual</div>
            <div style="font-size: 18px; font-weight: 700; color: #F8FAFC; margin-top: 4px;">${formatCOP(cuotaConfig)}</div>
          </div>
          <div style="background: rgba(255,255,255,0.03); border-left: 3px solid #10B981; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: var(--text2, #94A3B8);">Utilidad Acumulada</div>
            <div style="font-size: 18px; font-weight: 700; color: #10B981; margin-top: 4px;">${formatCOP(avanceUtilidad)}</div>
          </div>
          <div style="background: rgba(255,255,255,0.03); border-left: 3px solid #F59E0B; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: var(--text2, #94A3B8);">Facturación Bruta</div>
            <div style="font-size: 18px; font-weight: 700; color: #F59E0B; margin-top: 4px;">${formatCOP(mercancia)}</div>
          </div>
        </div>

        <div style="background: rgba(255,255,255,0.05); height: 8px; border-radius: 4px; overflow: hidden;">
          <div style="background: linear-gradient(90deg, #3B82F6, #10B981); height: 100%; width: ${Math.min(pctCumplimiento, 100)}%;"></div>
        </div>
      </div>
    `;
  }

  // B. VISTA DIRECTOR DE GRUPO
  function renderDirectorAvance(user, ventas, structure) {
    const container = document.getElementById("director-avance-panel");
    if (!container) return;

    const directorGrupo = user.directorGroup || user.group || "Grupo Beltran";
    
    // Filtrar ventas del grupo
    const ventasGrupo = ventas.filter(v => (v.Grupo_Personal || "").toLowerCase().includes(directorGrupo.toLowerCase().replace("grupo ", "")));
    const totalUtilidadGrupo = ventasGrupo.reduce((acc, v) => acc + (Number(v.Valor_Utilidad) || 0), 0);
    const totalMercanciaGrupo = ventasGrupo.reduce((acc, v) => acc + (Number(v.Valor_Mercancia) || 0), 0);

    let rowsHtml = ventasGrupo.map(v => `
      <tr style="border-bottom: 1px solid rgba(255,255,255,0.05);">
        <td style="padding: 10px; font-weight: 600; color: var(--text);">${v.Descripcion}</td>
        <td style="padding: 10px; text-align: right; color: var(--text2);">${formatCOP(v.Valor_Mercancia)}</td>
        <td style="padding: 10px; text-align: right; color: #10B981; font-weight: 700;">${formatCOP(v.Valor_Utilidad)}</td>
        <td style="padding: 10px; text-align: center; font-weight: 600; color: #38BDF8;">${v.Porcentaje_Utilidad}%</td>
      </tr>
    `).join("");

    container.innerHTML = `
      <div style="background: var(--card, #1E293B); border: 1px solid var(--border, rgba(255,255,255,0.1)); border-radius: 12px; padding: 20px; margin-bottom: 24px;">
        <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: var(--corp-cyan, #38BDF8);">Avance Diario del Equipo • ${directorGrupo}</div>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin: 16px 0;">
          <div style="background: rgba(255,255,255,0.03); border-left: 3px solid #10B981; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: var(--text2);">Utilidad Acumulada Equipo</div>
            <div style="font-size: 20px; font-weight: 700; color: #10B981; margin-top: 4px;">${formatCOP(totalUtilidadGrupo)}</div>
          </div>
          <div style="background: rgba(255,255,255,0.03); border-left: 3px solid #3B82F6; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: var(--text2);">Ventas Totales Equipo</div>
            <div style="font-size: 20px; font-weight: 700; color: #3B82F6; margin-top: 4px;">${formatCOP(totalMercanciaGrupo)}</div>
          </div>
        </div>

        <div style="overflow-x: auto;">
          <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
            <thead>
              <tr style="background: rgba(255,255,255,0.04); color: var(--text2); text-transform: uppercase;">
                <th style="padding: 8px; text-align: left;">Ejecutivo</th>
                <th style="padding: 8px; text-align: right;">Ventas</th>
                <th style="padding: 8px; text-align: right;">Utilidad</th>
                <th style="padding: 8px; text-align: center;">% Margen</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml || "<tr><td colspan='4' style='padding:12px;text-align:center;color:var(--text2)'>Sin ventas registradas hoy</td></tr>"}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  // C. VISTA GERENCIA GENERAL
  function renderGerenciaAvance(user, ventas, structure) {
    const container = document.getElementById("gerencia-avance-panel");
    if (!container) return;

    const totalEmpresaUtilidad = ventas.reduce((acc, v) => acc + (Number(v.Valor_Utilidad) || 0), 0);
    const totalEmpresaMercancia = ventas.reduce((acc, v) => acc + (Number(v.Valor_Mercancia) || 0), 0);

    container.innerHTML = `
      <div style="background: var(--card, #1E293B); border: 1px solid var(--border, rgba(255,255,255,0.1)); border-radius: 12px; padding: 20px; margin-bottom: 24px;">
        <div style="display: flex; justify-content: space-between; align-items: center;">
          <div>
            <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #A855F7;">Resumen de Utilidad Diario • Gerencia</div>
            <h3 style="margin: 4px 0 0; font-size: 20px; color: var(--text);">Consolidado Mes en Curso</h3>
          </div>
          <div style="text-align: right;">
            <div style="font-size: 11px; color: var(--text2);">Utilidad Total Empresa</div>
            <div style="font-size: 22px; font-weight: 800; color: #10B981;">${formatCOP(totalEmpresaUtilidad)}</div>
          </div>
        </div>
      </div>
    `;
  }

  // Exponer al ámbito global
  window.initAvanceDiario = initAvanceDiario;

  // Auto-inicializar cuando el DOM esté listo
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initAvanceDiario);
  } else {
    initAvanceDiario();
  }
})();
