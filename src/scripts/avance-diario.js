// ════════════════════════════════════════════════════════════════════════
//  FORECAST 2026 - MÓDULO DIARIO DE CUOTA Y AVANCE COMERCIAL (100% INTEGRADO)
// ════════════════════════════════════════════════════════════════════════

(function () {
  const UTILIDAD_API_AUTH_URL = "http://152.200.146.226:50010/api/getKey";
  const UTILIDAD_API_DATA_URL = "http://152.200.146.226:50010/consultas/api/consultaUtilidadComercialesDashboardPBI";
  const UTILIDAD_CREDS = { username: "powerbi", password: "3xpress#2025" };

  let cachedVentasData = null;

  function formatCOP(val) {
    if (val === null || val === undefined || isNaN(val)) return "$0 COP";
    return new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(val);
  }

  function getBogotaDateParts() {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "America/Bogota",
      year: "numeric",
      month: "2-digit",
      day: "2-digit"
    }).formatToParts(new Date()).reduce((parts, part) => {
      if (part.type !== "literal") parts[part.type] = part.value;
      return parts;
    }, {});
  }

  function getMonthDateRange() {
    const { year, month, day } = getBogotaDateParts();

    return {
      fechaInicial: `${year}-${month}-01`,
      fechaFinal: `${year}-${month}-${day}`
    };
  }

  function isCurrentMonthRow(row) {
    const currentMonth = getMonthDateRange().fechaInicial.slice(0, 7);
    if (typeof window.getMonth === "function" && typeof window.getRowDateValue === "function") {
      return window.getMonth(window.getRowDateValue(row)) === currentMonth;
    }
    const raw = row && (row["FECHA DIA/MES/AÑO"] || row.FECHA);
    if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
      const parts = new Intl.DateTimeFormat("en-CA", {
        timeZone: "America/Bogota",
        year: "numeric",
        month: "2-digit"
      }).formatToParts(raw).reduce((acc, part) => {
        if (part.type !== "literal") acc[part.type] = part.value;
        return acc;
      }, {});
      return `${parts.year}-${parts.month}` === currentMonth;
    }
    const text = String(raw || "").trim();
    const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})/);
    if (isoMatch) return `${isoMatch[1]}-${isoMatch[2].padStart(2, "0")}` === currentMonth;
    const localMatch = text.match(/^\d{1,2}[/-](\d{1,2})[/-](\d{4})/);
    return Boolean(localMatch && `${localMatch[2]}-${localMatch[1].padStart(2, "0")}` === currentMonth);
  }

  // Estructura oficial de cuotas y grupos 2026
  const ESTRUCTURA_DIRECTORES = {
    novoa: {
      nombre: "Rafael Novoa",
      vendedores: [
        { nombre: "Carolina Sánchez Pacheco", cuota: 18000000 },
        { nombre: "Rafael Francisco Novoa", cuota: 48000000 },
        { nombre: "Rosmira Rojas Puentes", cuota: 18000000 },
        { nombre: "Mario Reyes Gutierrez", cuota: 18000000 },
        { nombre: "Wilson Fernando Sánchez", cuota: 18000000 },
        { nombre: "María Eugenia Cruz", cuota: 18000000 },
        { nombre: "Javier Antonio Cortés", cuota: 18000000 },
        { nombre: "Rosa María Mendoza", cuota: 18000000 },
        { nombre: "Mariela Ramírez Castro", cuota: 18000000 },
        { nombre: "Jenny Alexandra González", cuota: 18000000 },
        { nombre: "María Paola Briceño", cuota: 48000000 },
        { nombre: "Jhonatan Camilo Hernández", cuota: 48000000 },
        { nombre: "Yeison Alonso Urrego", cuota: 48000000 },
        { nombre: "Jhonatan Steven Acevedo", cuota: 48000000 },
        { nombre: "Jasbleidy Johana Mojica", cuota: 48000000 },
        { nombre: "Diana Catalina Castro", cuota: 48000000 },
        { nombre: "Dafne Lizeth Ruiz Beltrán", cuota: 48000000 }
      ]
    },
    caballero: {
      nombre: "Angélica Caballero",
      vendedores: [
        { nombre: "Ángela Torres", cuota: 18000000 },
        { nombre: "Yurany Andrea Vargas", cuota: 18000000 },
        { nombre: "Alejandra Velásquez", cuota: 18000000 },
        { nombre: "Fernando Quiñonez", cuota: 18000000 },
        { nombre: "Jasbleidy Mójica", cuota: 48000000 },
        { nombre: "Johanna Jaime", cuota: 18000000 },
        { nombre: "Dayana Chala", cuota: 18000000 },
        { nombre: "Yovanny Herrera", cuota: 18000000 },
        { nombre: "César Céspedes", cuota: 28000000 },
        { nombre: "Daniel Galindo", cuota: 28000000 },
        { nombre: "Adriana Cucaita", cuota: 18000000 }
      ]
    },
    beltran: {
      nombre: "Óscar Beltrán",
      vendedores: [
        { nombre: "Dilma Constanza Cuesta", cuota: 18000000 },
        { nombre: "Claudia Patricia Triana", cuota: 18000000 },
        { nombre: "Angie Tatiana Parra", cuota: 18000000 },
        { nombre: "Leidy Astrid Jiménez Murcia", cuota: 18000000 },
        { nombre: "Óscar Alejandro Beltrán", cuota: 48000000 },
        { nombre: "Johanna Jaime Murcia", cuota: 18000000 },
        { nombre: "Julieth Milena Galindo Fino", cuota: 18000000 },
        { nombre: "Karent Carrillo Marin", cuota: 18000000 }
      ]
    },
    romero: {
      nombre: "Miller Romero",
      vendedores: [
        { nombre: "Freddy Andrés Peña Sánchez", cuota: 28000000 },
        { nombre: "Dayana Marcela Chalá", cuota: 18000000 },
        { nombre: "Daniel Galindo Girón", cuota: 28000000 },
        { nombre: "Lington Linares Linares", cuota: 18000000 },
        { nombre: "Jair Yovanny Herrera", cuota: 18000000 }
      ]
    }
  };

  function normText(str) {
    return String(str || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
  }

  function resolveDirectorKey(dirName) {
    const norm = normText(dirName);
    if (norm.includes("novoa") || norm.includes("rafael")) return "novoa";
    if (norm.includes("caballero") || norm.includes("angelica")) return "caballero";
    if (norm.includes("beltran") || norm.includes("oscar")) return "beltran";
    if (norm.includes("romero") || norm.includes("miller")) return "romero";
    return "novoa";
  }

  async function fetchApiVentas() {
    if (cachedVentasData) return cachedVentasData;

    try {
      const { fechaInicial, fechaFinal } = getMonthDateRange();
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
          return cachedVentasData;
        }
      }
    } catch (e) {
      console.warn("[AVANCE DIARIO] Petición API bloqueada por CORS en navegador. Calculando desde dataset de Forecast.", e.message);
    }
    return [];
  }

  // 1. VISTA DIRECTOR
  window.renderAvanceDiarioForDirector = async function (directorName) {
    const container = document.getElementById("director-avance-panel");
    if (!container) return;

    const dirKey = resolveDirectorKey(directorName);
    const dirInfo = ESTRUCTURA_DIRECTORES[dirKey] || ESTRUCTURA_DIRECTORES.novoa;
    const vendedores = dirInfo.vendedores;

    // Cuota Total del Grupo
    const totalCuotaGrupo = vendedores.reduce((acc, v) => acc + v.cuota, 0);

    // Obtener utilidad y mercancia de la API o del Forecast ALL_DATA
    let totalUtilidadGrupo = 0;
    let totalMercanciaGrupo = 0;

    const apiVentas = await fetchApiVentas();

    if (apiVentas && apiVentas.length) {
      vendedores.forEach(v => {
        const vNorm = normText(v.nombre);
        const match = apiVentas.find(item => normText(item.Descripcion).includes(vNorm.split(" ")[0]));
        if (match) {
          totalUtilidadGrupo += Number(match.Valor_Utilidad) || 0;
          totalMercanciaGrupo += Number(match.Valor_Mercancia) || 0;
        }
      });
    }

    // Si la API no respondió por CORS, extraer de window.ALL_DATA de Forecast
    if (totalUtilidadGrupo === 0 && window.ALL_DATA && window.ALL_DATA.length) {
      const dataDir = window.ALL_DATA.filter(r =>
        normText(r['DIRECTOR']).includes(dirKey) && isCurrentMonthRow(r)
      );
      totalUtilidadGrupo = dataDir.reduce((sum, r) => {
        const utilObj = typeof window.getUtilidad === 'function' ? window.getUtilidad(r) : { valor: 0 };
        return sum + (utilObj.valor || 0);
      }, 0);

      totalMercanciaGrupo = dataDir.reduce((sum, r) => {
        const copVal = typeof window.toCOP === 'function' ? window.toCOP(r) : 0;
        return sum + copVal;
      }, 0);
    }

    const pctAvance = totalCuotaGrupo > 0 ? (totalUtilidadGrupo / totalCuotaGrupo) * 100 : 0;
    const faltante = Math.max(0, totalCuotaGrupo - totalUtilidadGrupo);

    const pctAvanceLimitado = Math.min(Math.max(pctAvance, 0), 100);

    container.innerHTML = `
      <section class="quota-summary quota-summary--director" aria-label="Avance de cuota y utilidad de ${dirInfo.nombre}">
        <div class="quota-summary__header">
          <div>
            <div class="quota-summary__eyebrow"><span class="quota-summary__signal" aria-hidden="true"></span>Avance diario de cuota y utilidad</div>
            <div class="quota-summary__title">${dirInfo.nombre}</div>
            <div class="quota-summary__subtitle">Seguimiento acumulado vs Cuota del Grupo (${vendedores.length} ejecutivos asignados)</div>
          </div>
          <div class="quota-summary__status">
            <strong>${pctAvance.toFixed(1)}%</strong>
            <span>Cumplimiento</span>
          </div>
        </div>

        <div class="quota-summary__metrics">
          <article class="quota-metric quota-metric--quota">
            <span class="quota-metric__icon" aria-hidden="true">◎</span>
            <div>
              <span>Cuota Mensual Grupo</span>
              <strong>${formatCOP(totalCuotaGrupo)}</strong>
              <small>Meta asignada para ${vendedores.length} ejecutivos</small>
            </div>
          </article>

          <article class="quota-metric quota-metric--profit">
            <span class="quota-metric__icon" aria-hidden="true">↗</span>
            <div>
              <span>Utilidad Lograda (Avance)</span>
              <strong>${formatCOP(totalUtilidadGrupo)}</strong>
              <small>Margen acumulado en el mes</small>
            </div>
          </article>

          <article class="quota-metric quota-metric--remaining">
            <span class="quota-metric__icon" aria-hidden="true">△</span>
            <div>
              <span>Faltante para la Meta</span>
              <strong>${formatCOP(faltante)}</strong>
              <small>Diferencia sobre cuota</small>
            </div>
          </article>
        </div>

        <div class="quota-summary__progress-row">
          <div class="quota-summary__progress" role="progressbar" aria-label="Cumplimiento de cuota del grupo" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${pctAvanceLimitado.toFixed(1)}">
            <span style="--progress:${pctAvanceLimitado}%"></span>
          </div>
          <strong>${pctAvance.toFixed(1)}%</strong>
        </div>
      </section>
    `;
  };

  // 2. VISTA EJECUTIVO
  window.renderAvanceDiarioForEjecutivo = async function (execName) {
    const container = document.getElementById("ejecutivo-avance-panel");
    if (!container) return;

    const eNorm = normText(execName);
    let cuota = 18000000;

    Object.values(ESTRUCTURA_DIRECTORES).forEach(dir => {
      dir.vendedores.forEach(v => {
        if (normText(v.nombre).includes(eNorm.split(" ")[0])) {
          cuota = v.cuota;
        }
      });
    });

    let avanceUtilidad = 0;
    let mercancia = 0;

    const apiVentas = await fetchApiVentas();
    if (apiVentas && apiVentas.length) {
      const match = apiVentas.find(item => normText(item.Descripcion).includes(eNorm.split(" ")[0]));
      if (match) {
        avanceUtilidad = Number(match.Valor_Utilidad) || 0;
        mercancia = Number(match.Valor_Mercancia) || 0;
      }
    }

    if (avanceUtilidad === 0 && window.ALL_DATA && window.ALL_DATA.length) {
      const dataEj = window.ALL_DATA.filter(r =>
        normText(r['COMERCIAL']).includes(eNorm.split(" ")[0]) && isCurrentMonthRow(r)
      );
      avanceUtilidad = dataEj.reduce((sum, r) => {
        const utilObj = typeof window.getUtilidad === 'function' ? window.getUtilidad(r) : { valor: 0 };
        return sum + (utilObj.valor || 0);
      }, 0);

      mercancia = dataEj.reduce((sum, r) => {
        const copVal = typeof window.toCOP === 'function' ? window.toCOP(r) : 0;
        return sum + copVal;
      }, 0);
    }

    const pctAvance = cuota > 0 ? (avanceUtilidad / cuota) * 100 : 0;

    const pctAvanceLimitado = Math.min(Math.max(pctAvance, 0), 100);

    container.innerHTML = `
      <section class="quota-summary quota-summary--executive" aria-label="Cuota y utilidad de ${execName || 'Ejecutivo Comercial'}">
        <div class="quota-summary__header">
          <div>
            <div class="quota-summary__eyebrow"><span class="quota-summary__signal" aria-hidden="true"></span>Mi cuota y avance diario comercial</div>
            <div class="quota-summary__title">${execName || 'Ejecutivo Comercial'}</div>
            <div class="quota-summary__subtitle">Utilidad acumulada frente a la meta mensual</div>
          </div>
          <div class="quota-summary__status">
            <strong>${pctAvance.toFixed(1)}%</strong>
            <span>De cuota alcanzada</span>
          </div>
        </div>

        <div class="quota-summary__metrics">
          <article class="quota-metric quota-metric--quota">
            <span class="quota-metric__icon" aria-hidden="true">◎</span>
            <div>
              <span>Mi Cuota Mensual</span>
              <strong>${formatCOP(cuota)}</strong>
              <small>Meta comercial asignada</small>
            </div>
          </article>

          <article class="quota-metric quota-metric--profit">
            <span class="quota-metric__icon" aria-hidden="true">↗</span>
            <div>
              <span>Utilidad Lograda (Avance)</span>
              <strong>${formatCOP(avanceUtilidad)}</strong>
              <small>Margen acumulado en el mes</small>
            </div>
          </article>

          <article class="quota-metric quota-metric--sales">
            <span class="quota-metric__icon" aria-hidden="true">$</span>
            <div>
              <span>Ventas Totales Brutas</span>
              <strong>${formatCOP(mercancia)}</strong>
              <small>Mercancía acumulada en el mes</small>
            </div>
          </article>
        </div>

        <div class="quota-summary__progress-row">
          <div class="quota-summary__progress" role="progressbar" aria-label="Cumplimiento de cuota del ejecutivo" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${pctAvanceLimitado.toFixed(1)}">
            <span style="--progress:${pctAvanceLimitado}%"></span>
          </div>
          <strong>${pctAvance.toFixed(1)}%</strong>
        </div>
      </section>
    `;
  };

})();
