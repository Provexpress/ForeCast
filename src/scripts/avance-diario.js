// ════════════════════════════════════════════════════════════════════════
//  FORECAST 2026 - MÓDULO DIARIO DE CUOTA Y AVANCE COMERCIAL (INTEGRADO)
// ════════════════════════════════════════════════════════════════════════

(function () {
  const UTILIDAD_API_AUTH_URL = "http://152.200.146.226:50010/api/getKey";
  const UTILIDAD_API_DATA_URL = "http://152.200.146.226:50010/consultas/api/consultaUtilidadComercialesDashboardPBI";
  const UTILIDAD_CREDS = { username: "powerbi", password: "3xpress#2025" };

  let cachedVentasData = null;
  let isFetching = false;

  // Formateadores de moneda y números
  function formatCOP(val) {
    if (val === null || val === undefined || isNaN(val)) return "$0 COP";
    return new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(val);
  }

  function formatAbr(num) {
    if (!num) return "$0";
    if (Math.abs(num) >= 1e9) return "$" + (num / 1e9).toFixed(2) + " Bill";
    if (Math.abs(num) >= 1e6) return "$" + (num / 1e6).toFixed(2) + " MM";
    if (Math.abs(num) >= 1e3) return "$" + (num / 1e3).toFixed(1) + " K";
    return "$" + Math.round(num);
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

  // Mapa estático de cuotas por vendedor
  const CUOTAS_VENDEDORES = {
    "carolina sanchez": 18000000,
    "rafael francisco novoa": 48000000,
    "dilma constanza cuesta": 18000000,
    "claudia patricia triana": 18000000,
    "maria paola briceño": 48000000,
    "jhonatan camilo hernandez": 48000000,
    "yeison alonso urrego": 48000000,
    "jhonatan steven acevedo": 48000000,
    "jasbleidy johana mojica": 48000000,
    "freddy andres peña": 28000000,
    "maria alejandra velasquez": 18000000,
    "angie tatiana parra": 18000000,
    "leidy astrid jimenez": 18000000,
    "oscar alejandro beltran": 48000000,
    "johanna jaime murcia": 18000000,
    "rosa maria mendoza": 18000000,
    "fernando alberto quiñonez": 18000000,
    "dayana marcela chala": 18000000,
    "javier antonio cortes": 18000000,
    "maria eugenia cruz": 18000000,
    "karent carrillo marin": 18000000,
    "daniel galindo giron": 28000000,
    "maria angelica caballero": 48000000,
    "rosmira rojas puentes": 18000000,
    "diana catalina castro": 48000000,
    "mariela ramírez castro": 18000000,
    "mario reyes gutierrez": 18000000,
    "lington linares linares": 18000000,
    "julieth milena galindo": 18000000,
    "dafne lizeth ruiz": 48000000,
    "wilson fernando sánchez": 18000000,
    "cesar augusto cespedes": 28000000,
    "gina paola garcia": 18000000,
    "jair yovanny herrea": 18000000,
    "jessica lorena valencia": 18000000,
    "maria angelica alvarado": 18000000,
    "angela rocio torres": 18000000,
    "yurany andrea vargas": 18000000,
    "jenny alexandra gonzalez": 18000000,
    "juan david martínez": 14000000
  };

  async function getVentasDiarias() {
    if (cachedVentasData) return cachedVentasData;
    if (isFetching) return [];

    isFetching = true;
    try {
      const { fechaInicial, fechaFinal } = getMonthDateRange();
      
      // Intentar fetch directo
      const authRes = await fetch(UTILIDAD_API_AUTH_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(UTILIDAD_CREDS)
      });

      if (authRes.ok) {
        const authData = await authRes.json();
        const token = authData.token || authData.access_token;

        const dataRes = await fetch(UTILIDAD_API_DATA_URL, {
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

        if (dataRes.ok) {
          const json = await dataRes.json();
          cachedVentasData = json.response || [];
          isFetching = false;
          return cachedVentasData;
        }
      }
    } catch (e) {
      console.warn("[AVANCE DIARIO] Petición API no permitida por navegador (CORS/Mixed Content). Usando cálculo interno de cuotas y avance de negocios.", e.message);
    }

    isFetching = false;
    return [];
  }

  // A. VISTA DIRECTOR
  window.renderAvanceDiarioForDirector = async function (directorName) {
    const container = document.getElementById("director-avance-panel");
    if (!container) return;

    const dirNorm = (directorName || "").toLowerCase().trim();
    const ventas = await getVentasDiarias();

    // Obtener datos del director desde la estructura o data global
    let totalCuotaGrupo = 0;
    let totalUtilidadGrupo = 0;
    let totalMercanciaGrupo = 0;

    // Calcular cuotas del grupo según vendedores de ese director
    const estructura = window.FORECAST_STRUCTURE || {};
    let ejecutivos = [];
    if (estructura.getEjecutivosByDirector) {
      ejecutivos = estructura.getEjecutivosByDirector(directorName) || [];
    }

    if (!ejecutivos.length && window.ALL_DATA) {
      ejecutivos = [...new Set(window.ALL_DATA.filter(r => (r['DIRECTOR'] || '').trim().toLowerCase().includes(dirNorm.replace("grupo ", ""))).map(r => r['COMERCIAL']))];
    }

    ejecutivos.forEach(e => {
      const eNorm = (e || "").toLowerCase().trim();
      let cuota = 18000000;
      for (const [key, val] of Object.entries(CUOTAS_VENDEDORES)) {
        if (eNorm.includes(key) || key.includes(eNorm)) {
          cuota = val; break;
        }
      }
      totalCuotaGrupo += cuota;

      // Sumar utilidad de API si está disponible
      const venta = ventas.find(v => (v.Descripcion || "").toLowerCase().includes(eNorm.split(" ")[0]));
      if (venta) {
        totalUtilidadGrupo += Number(venta.Valor_Utilidad) || 0;
        totalMercanciaGrupo += Number(venta.Valor_Mercancia) || 0;
      }
    });

    // Si la API no retornó datos por CORS, calcular desde los datos cargados de negocios ganados del mes
    if (totalUtilidadGrupo === 0 && window.ALL_DATA) {
      const dataDir = window.ALL_DATA.filter(r => (r['DIRECTOR'] || '').trim().toLowerCase().includes(dirNorm.replace("grupo ", "")));
      totalUtilidadGrupo = dataDir.filter(r => r['ESTADO'] === 'GANADA').reduce((sum, r) => sum + (Number(r['UTILIDAD COP']) || Number(r['MONTO VENTA CLIENTE']) * 0.15 || 0), 0);
      totalMercanciaGrupo = dataDir.reduce((sum, r) => sum + (Number(r['MONTO VENTA CLIENTE']) || 0), 0);
    }

    const pctAvance = totalCuotaGrupo > 0 ? (totalUtilidadGrupo / totalCuotaGrupo) * 100 : 0;

    container.innerHTML = `
      <div style="background: linear-gradient(135deg, rgba(30,41,59,0.9), rgba(15,23,42,0.95)); border: 1px solid rgba(56,189,248,0.25); border-radius: 12px; padding: 20px; margin: 16px 0 24px; box-shadow: 0 4px 16px rgba(0,0,0,0.2);">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; flex-wrap: wrap; gap: 8px;">
          <div>
            <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px; color: #38BDF8;">AVANCE DIARIO DE CUOTA Y UTILIDAD • ${directorName || 'DIRECTOR'}</div>
            <div style="font-size: 13px; color: #94A3B8; margin-top: 2px;">Seguimiento acumulado del mes en curso vs Cuota del Grupo</div>
          </div>
          <div style="background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.3); color: #10B981; font-weight: 800; font-size: 13px; padding: 6px 14px; border-radius: 20px;">
            ${pctAvance.toFixed(1)}% CUMPLIMIENTO
          </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 14px;">
          <div style="background: rgba(255,255,255,0.03); border-left: 4px solid #3B82F6; padding: 12px 14px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Cuota Mensual Grupo</div>
            <div style="font-size: 20px; font-weight: 800; color: #F8FAFC; margin-top: 4px;">${formatCOP(totalCuotaGrupo)}</div>
            <div style="font-size: 10px; color: #64748B; margin-top: 2px;">${ejecutivos.length} comerciales asignados</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border-left: 4px solid #10B981; padding: 12px 14px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Utilidad Lograda (Avance)</div>
            <div style="font-size: 20px; font-weight: 800; color: #10B981; margin-top: 4px;">${formatCOP(totalUtilidadGrupo)}</div>
            <div style="font-size: 10px; color: #64748B; margin-top: 2px;">Margen de ganancias del periodo</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border-left: 4px solid #F59E0B; padding: 12px 14px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Faltante para la Meta</div>
            <div style="font-size: 20px; font-weight: 800; color: #F59E0B; margin-top: 4px;">${formatCOP(Math.max(0, totalCuotaGrupo - totalUtilidadGrupo))}</div>
            <div style="font-size: 10px; color: #64748B; margin-top: 2px;">Diferencia sobre cuota</div>
          </div>
        </div>

        <div style="background: rgba(255,255,255,0.08); height: 10px; border-radius: 5px; overflow: hidden;">
          <div style="background: linear-gradient(90deg, #3B82F6, #10B981); height: 100%; width: ${Math.min(pctAvance, 100)}%;"></div>
        </div>
      </div>
    `;
  };

  // B. VISTA EJECUTIVO
  window.renderAvanceDiarioForEjecutivo = async function (execName) {
    const container = document.getElementById("ejecutivo-avance-panel");
    if (!container) return;

    const eNorm = (execName || "").toLowerCase().trim();
    const ventas = await getVentasDiarias();

    let cuota = 18000000;
    for (const [key, val] of Object.entries(CUOTAS_VENDEDORES)) {
      if (eNorm.includes(key) || key.includes(eNorm)) {
        cuota = val; break;
      }
    }

    let avanceUtilidad = 0;
    let mercancia = 0;

    const venta = ventas.find(v => (v.Descripcion || "").toLowerCase().includes(eNorm.split(" ")[0]));
    if (venta) {
      avanceUtilidad = Number(venta.Valor_Utilidad) || 0;
      mercancia = Number(venta.Valor_Mercancia) || 0;
    } else if (window.ALL_DATA) {
      const dataEj = window.ALL_DATA.filter(r => (r['COMERCIAL'] || '').trim().toLowerCase().includes(eNorm.split(" ")[0]));
      avanceUtilidad = dataEj.filter(r => r['ESTADO'] === 'GANADA').reduce((sum, r) => sum + (Number(r['UTILIDAD COP']) || Number(r['MONTO VENTA CLIENTE']) * 0.15 || 0), 0);
      mercancia = dataEj.reduce((sum, r) => sum + (Number(r['MONTO VENTA CLIENTE']) || 0), 0);
    }

    const pctAvance = cuota > 0 ? (avanceUtilidad / cuota) * 100 : 0;

    container.innerHTML = `
      <div style="background: linear-gradient(135deg, rgba(30,41,59,0.9), rgba(15,23,42,0.95)); border: 1px solid rgba(168,85,247,0.3); border-radius: 12px; padding: 20px; margin: 16px 0 24px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; flex-wrap: wrap; gap: 8px;">
          <div>
            <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px; color: #C084FC;">MI CUOTA Y AVANCE DIARIO COMERCIAL</div>
            <div style="font-size: 16px; font-weight: 700; color: #F8FAFC; margin-top: 2px;">${execName || 'Ejecutivo Comercial'}</div>
          </div>
          <div style="background: rgba(168,85,247,0.15); border: 1px solid rgba(168,85,247,0.4); color: #C084FC; font-weight: 800; font-size: 13px; padding: 6px 14px; border-radius: 20px;">
            ${pctAvance.toFixed(1)}% DE CUOTA ALCANZADA
          </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 14px;">
          <div style="background: rgba(255,255,255,0.03); border-left: 4px solid #8B5CF6; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Mi Cuota Mensual</div>
            <div style="font-size: 20px; font-weight: 800; color: #F8FAFC; margin-top: 4px;">${formatCOP(cuota)}</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border-left: 4px solid #10B981; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Utilidad Lograda (Avance)</div>
            <div style="font-size: 20px; font-weight: 800; color: #10B981; margin-top: 4px;">${formatCOP(avanceUtilidad)}</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border-left: 4px solid #3B82F6; padding: 12px; border-radius: 6px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Ventas Totales Brutas</div>
            <div style="font-size: 20px; font-weight: 800; color: #3B82F6; margin-top: 4px;">${formatCOP(mercancia)}</div>
          </div>
        </div>

        <div style="background: rgba(255,255,255,0.08); height: 10px; border-radius: 5px; overflow: hidden;">
          <div style="background: linear-gradient(90deg, #8B5CF6, #10B981); height: 100%; width: ${Math.min(pctAvance, 100)}%;"></div>
        </div>
      </div>
    `;
  };

  // C. VISTA GERENCIA
  window.renderAvanceDiarioForGerencia = async function () {
    const container = document.getElementById("gerencia-avance-panel");
    if (!container) return;

    let totalCuotaEmpresa = 878000000; // Total cuotas 2026
    let totalUtilidadEmpresa = 0;

    const ventas = await getVentasDiarias();
    if (ventas.length) {
      totalUtilidadEmpresa = ventas.reduce((acc, v) => acc + (Number(v.Valor_Utilidad) || 0), 0);
    } else if (window.ALL_DATA) {
      totalUtilidadEmpresa = window.ALL_DATA.filter(r => r['ESTADO'] === 'GANADA').reduce((sum, r) => sum + (Number(r['UTILIDAD COP']) || Number(r['MONTO VENTA CLIENTE']) * 0.15 || 0), 0);
    }

    const pctAvance = (totalUtilidadEmpresa / totalCuotaEmpresa) * 100;

    container.innerHTML = `
      <div style="background: linear-gradient(135deg, rgba(30,41,59,0.9), rgba(15,23,42,0.95)); border: 1px solid rgba(16,185,129,0.3); border-radius: 12px; padding: 20px; margin: 16px 0 24px;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
          <div>
            <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px; color: #10B981;">RESUMEN GERENCIAL DE CUOTA Y AVANCE DIARIO</div>
            <div style="font-size: 18px; font-weight: 800; color: #F8FAFC; margin-top: 2px;">Utilidad Total Empresa vs Meta Acumulada</div>
          </div>
          <div style="background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.4); color: #10B981; font-weight: 800; font-size: 14px; padding: 6px 16px; border-radius: 20px;">
            ${formatCOP(totalUtilidadEmpresa)} (${pctAvance.toFixed(1)}% Meta)
          </div>
        </div>
      </div>
    `;
  };

  // Escuchar cambios de selector o renderizado de páginas
  function hookIntoForecastRender() {
    // Hook en renderDirector
    if (window.renderDirector) {
      const origRenderDir = window.renderDirector;
      window.renderDirector = function (...args) {
        origRenderDir.apply(this, args);
        const selDir = document.getElementById("sel-director");
        const dirName = selDir ? selDir.value : "";
        window.renderAvanceDiarioForDirector(dirName);
      };
    }

    // Hook en renderEjecutivo
    if (window.renderEjecutivo) {
      const origRenderEj = window.renderEjecutivo;
      window.renderEjecutivo = function (...args) {
        origRenderEj.apply(this, args);
        const selEj = document.getElementById("sel-ejecutivo");
        const ejName = selEj ? selEj.value : "";
        window.renderAvanceDiarioForEjecutivo(ejName);
      };
    }
  }

  // Auto-iniciar al cargar la página
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      hookIntoForecastRender();
    });
  } else {
    hookIntoForecastRender();
  }
})();
