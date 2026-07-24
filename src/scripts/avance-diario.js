// ════════════════════════════════════════════════════════════════════════
//  FORECAST 2026 - MÓDULO DIARIO DE CUOTA Y AVANCE COMERCIAL (ROBUSTO)
// ════════════════════════════════════════════════════════════════════════

(function () {
  const UTILIDAD_API_AUTH_URL = "http://152.200.146.226:50010/api/getKey";
  const UTILIDAD_API_DATA_URL = "http://152.200.146.226:50010/consultas/api/consultaUtilidadComercialesDashboardPBI";
  const UTILIDAD_CREDS = { username: "powerbi", password: "3xpress#2025" };

  let cachedVentasData = null;

  // Formateadores de moneda y números
  function formatCOP(val) {
    if (val === null || val === undefined || isNaN(val)) return "$0 COP";
    return new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(val);
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

  // Estructura oficial de cuotas y grupos 2026
  const ESTRUCTURA_DIRECTORES = {
    novoa: {
      nombre: "Rafael Novoa",
      grupoApi: "Grupo Novoa",
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
      grupoApi: "Grupo Caballero",
      vendedores: [
        { nombre: "Ángela Rocío Torres", cuota: 18000000 },
        { nombre: "Yurany Andrea Vargas", cuota: 18000000 },
        { nombre: "María Alejandra Velásquez", cuota: 18000000 },
        { nombre: "Fernando Alberto Quiñonez", cuota: 18000000 },
        { nombre: "María Angélica Caballero", cuota: 48000000 },
        { nombre: "César Augusto Céspedes", cuota: 28000000 },
        { nombre: "Gina Paola García Quintero", cuota: 18000000 },
        { nombre: "Jessica Lorena Valencia", cuota: 18000000 },
        { nombre: "María Angélica Alvarado", cuota: 18000000 },
        { nombre: "Juan David Martínez", cuota: 14000000 }
      ]
    },
    beltran: {
      nombre: "Óscar Beltrán",
      grupoApi: "Grupo Beltran",
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
      grupoApi: "Grupo Romero",
      vendedores: [
        { nombre: "Freddy Andrés Peña Sánchez", cuota: 28000000 },
        { nombre: "Dayana Marcela Chalá", cuota: 18000000 },
        { nombre: "Daniel Galindo Girón", cuota: 28000000 },
        { nombre: "Lington Linares Linares", cuota: 18000000 },
        { nombre: "Jair Yovanny Herrera", cuota: 18000000 }
      ]
    }
  };

  // Normalización de texto para comparaciones
  function normText(str) {
    return String(str || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
  }

  // Identificar la clave del director
  function resolveDirectorKey(dirName) {
    const norm = normText(dirName);
    if (norm.includes("novoa") || norm.includes("rafael")) return "novoa";
    if (norm.includes("caballero") || norm.includes("angelica")) return "caballero";
    if (norm.includes("beltran") || norm.includes("oscar")) return "beltran";
    if (norm.includes("romero") || norm.includes("miller")) return "romero";
    return "novoa"; // Default
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
      console.warn("[AVANCE DIARIO] API no respondió directamente por políticas de navegador (CORS). Se calcula el avance acumulado comercial.", e.message);
    }
    return [];
  }

  // Renderizador principal para VISTA DIRECTOR
  window.renderAvanceDiarioForDirector = async function (directorName) {
    const container = document.getElementById("director-avance-panel");
    if (!container) return;

    const dirKey = resolveDirectorKey(directorName);
    const dirInfo = ESTRUCTURA_DIRECTORES[dirKey];
    const vendedores = dirInfo.vendedores;

    // 1. Cuota Total del Grupo
    const totalCuotaGrupo = vendedores.reduce((acc, v) => acc + v.cuota, 0);

    // 2. Intentar consultar ventas de la API o calcular desde datos del Forecast
    const apiVentas = await fetchApiVentas();
    let totalUtilidadGrupo = 0;
    let totalMercanciaGrupo = 0;

    vendedores.forEach(v => {
      const vNorm = normText(v.nombre);
      const match = apiVentas.find(item => normText(item.Descripcion).includes(vNorm.split(" ")[0]));
      if (match) {
        totalUtilidadGrupo += Number(match.Valor_Utilidad) || 0;
        totalMercanciaGrupo += Number(match.Valor_Mercancia) || 0;
      }
    });

    // Fallback: Si no hay respuesta API en el cliente por CORS, sumar desde ALL_DATA (negocios del mes)
    if (totalUtilidadGrupo === 0 && window.ALL_DATA && window.ALL_DATA.length) {
      const dataDir = window.ALL_DATA.filter(r => normText(r['DIRECTOR']).includes(dirKey));
      totalUtilidadGrupo = dataDir.filter(r => r['ESTADO'] === 'GANADA').reduce((sum, r) => sum + (Number(r['UTILIDAD COP']) || Number(r['MONTO VENTA CLIENTE']) * 0.15 || 0), 0);
      totalMercanciaGrupo = dataDir.reduce((sum, r) => sum + (Number(r['MONTO VENTA CLIENTE']) || 0), 0);
    }

    const pctAvance = totalCuotaGrupo > 0 ? (totalUtilidadGrupo / totalCuotaGrupo) * 100 : 0;
    const faltante = Math.max(0, totalCuotaGrupo - totalUtilidadGrupo);

    container.innerHTML = `
      <div style="background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid rgba(56,189,248,0.3); border-radius: 12px; padding: 22px; margin: 16px 0 24px; box-shadow: 0 8px 24px rgba(0,0,0,0.25); position: relative; overflow: hidden;">
        <div style="position: absolute; right: -20px; top: -20px; width: 140px; height: 140px; background: rgba(56,189,248,0.05); border-radius: 50%; pointer-events: none;"></div>
        
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-wrap: wrap; gap: 12px;">
          <div>
            <div style="font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.4px; color: #38BDF8; display: flex; align-items: center; gap: 6px;">
              <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #38BDF8;"></span>
              AVANCE DIARIO DE CUOTA Y UTILIDAD • ${dirInfo.nombre.toUpperCase()}
            </div>
            <div style="font-size: 13px; color: #94A3B8; margin-top: 4px;">Seguimiento acumulado del mes en curso vs Cuota del Grupo (${vendedores.length} ejecutivos asignados)</div>
          </div>
          
          <div style="background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.4); color: #34D399; font-weight: 800; font-size: 14px; padding: 8px 18px; border-radius: 20px; box-shadow: 0 2px 8px rgba(16,185,129,0.15);">
            ${pctAvance.toFixed(1)}% CUMPLIMIENTO
          </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 16px;">
          <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-left: 4px solid #3B82F6; padding: 14px 16px; border-radius: 8px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700; letter-spacing: 0.5px;">Cuota Mensual Grupo</div>
            <div style="font-size: 22px; font-weight: 800; color: #F8FAFC; margin-top: 4px;">${formatCOP(totalCuotaGrupo)}</div>
            <div style="font-size: 11px; color: #64748B; margin-top: 2px;">Meta asignada para ${vendedores.length} ejecutivos</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-left: 4px solid #10B981; padding: 14px 16px; border-radius: 8px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700; letter-spacing: 0.5px;">Utilidad Lograda (Avance)</div>
            <div style="font-size: 22px; font-weight: 800; color: #10B981; margin-top: 4px;">${formatCOP(totalUtilidadGrupo)}</div>
            <div style="font-size: 11px; color: #64748B; margin-top: 2px;">Margen acumulado en el mes</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-left: 4px solid #F59E0B; padding: 14px 16px; border-radius: 8px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700; letter-spacing: 0.5px;">Faltante para la Meta</div>
            <div style="font-size: 22px; font-weight: 800; color: #F59E0B; margin-top: 4px;">${formatCOP(faltante)}</div>
            <div style="font-size: 11px; color: #64748B; margin-top: 2px;">Brecha requerida para el 100%</div>
          </div>
        </div>

        <!-- Barra de progreso -->
        <div style="background: rgba(255,255,255,0.08); height: 10px; border-radius: 5px; overflow: hidden; position: relative;">
          <div style="background: linear-gradient(90deg, #3B82F6, #10B981); height: 100%; width: ${Math.min(pctAvance, 100)}%; transition: width 0.6s ease;"></div>
        </div>
      </div>
    `;
  };

  // Renderizador para VISTA EJECUTIVO
  window.renderAvanceDiarioForEjecutivo = async function (execName) {
    const container = document.getElementById("ejecutivo-avance-panel");
    if (!container) return;

    const eNorm = normText(execName);
    let cuota = 18000000; // Base por defecto

    // Buscar cuota del ejecutivo
    Object.values(ESTRUCTURA_DIRECTORES).forEach(dir => {
      dir.vendedores.forEach(v => {
        if (normText(v.nombre).includes(eNorm.split(" ")[0])) {
          cuota = v.cuota;
        }
      });
    });

    const apiVentas = await fetchApiVentas();
    let avanceUtilidad = 0;
    let mercancia = 0;

    const match = apiVentas.find(item => normText(item.Descripcion).includes(eNorm.split(" ")[0]));
    if (match) {
      avanceUtilidad = Number(match.Valor_Utilidad) || 0;
      mercancia = Number(match.Valor_Mercancia) || 0;
    } else if (window.ALL_DATA && window.ALL_DATA.length) {
      const dataEj = window.ALL_DATA.filter(r => normText(r['COMERCIAL']).includes(eNorm.split(" ")[0]));
      avanceUtilidad = dataEj.filter(r => r['ESTADO'] === 'GANADA').reduce((sum, r) => sum + (Number(r['UTILIDAD COP']) || Number(r['MONTO VENTA CLIENTE']) * 0.15 || 0), 0);
      mercancia = dataEj.reduce((sum, r) => sum + (Number(r['MONTO VENTA CLIENTE']) || 0), 0);
    }

    const pctAvance = cuota > 0 ? (avanceUtilidad / cuota) * 100 : 0;
    const faltante = Math.max(0, cuota - avanceUtilidad);

    container.innerHTML = `
      <div style="background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid rgba(192,132,252,0.3); border-radius: 12px; padding: 22px; margin: 16px 0 24px; box-shadow: 0 8px 24px rgba(0,0,0,0.25);">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-wrap: wrap; gap: 12px;">
          <div>
            <div style="font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 1.4px; color: #C084FC;">MI CUOTA Y AVANCE DIARIO COMERCIAL</div>
            <div style="font-size: 18px; font-weight: 800; color: #F8FAFC; margin-top: 2px;">${execName || 'Ejecutivo Comercial'}</div>
          </div>
          <div style="background: rgba(192,132,252,0.15); border: 1px solid rgba(192,132,252,0.4); color: #C084FC; font-weight: 800; font-size: 14px; padding: 8px 18px; border-radius: 20px;">
            ${pctAvance.toFixed(1)}% CUMPLIMIENTO DE CUOTA
          </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 16px;">
          <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-left: 4px solid #8B5CF6; padding: 14px 16px; border-radius: 8px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Mi Cuota Mensual</div>
            <div style="font-size: 22px; font-weight: 800; color: #F8FAFC; margin-top: 4px;">${formatCOP(cuota)}</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-left: 4px solid #10B981; padding: 14px 16px; border-radius: 8px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Utilidad Lograda (Avance)</div>
            <div style="font-size: 22px; font-weight: 800; color: #10B981; margin-top: 4px;">${formatCOP(avanceUtilidad)}</div>
          </div>

          <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-left: 4px solid #F59E0B; padding: 14px 16px; border-radius: 8px;">
            <div style="font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 700;">Faltante para la Meta</div>
            <div style="font-size: 22px; font-weight: 800; color: #F59E0B; margin-top: 4px;">${formatCOP(faltante)}</div>
          </div>
        </div>

        <div style="background: rgba(255,255,255,0.08); height: 10px; border-radius: 5px; overflow: hidden;">
          <div style="background: linear-gradient(90deg, #8B5CF6, #10B981); height: 100%; width: ${Math.min(pctAvance, 100)}%;"></div>
        </div>
      </div>
    `;
  };

})();
