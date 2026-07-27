// ============================================================================
// MÓDULO DE AVANCE DIARIO DE CUOTA Y UTILIDAD COMERCIAL (FORECAST 2026)
// Consume la API de Utilidad Comercial (o cache sincronizado src/data/api-utilidad-cache.json)
// ============================================================================

(function () {
  window.API_UTILIDAD_DATA = [];

  // 1. Cargar datos de la API de Utilidad en el cliente
  async function loadApiUtilidadCache() {
    try {
      const cachePath = 'src/data/api-utilidad-cache.json?v=' + Date.now();
      const resp = await fetch(cachePath);
      if (resp.ok) {
        const data = await resp.json();
        window.API_UTILIDAD_DATA = data.vendedores || [];
        console.log('✅ API Utilidad Comercial cargada en ForeCast:', window.API_UTILIDAD_DATA.length, 'registros de utilidad');
        
        // Disparar evento de actualización por si la vista ya se renderizó
        if (typeof window.refreshAvanceDiarioViews === 'function') {
          window.refreshAvanceDiarioViews();
        }
      }
    } catch (e) {
      console.warn('⚠️ No se pudo cargar api-utilidad-cache.json, usando fallback local.', e.message);
    }
  }

  // 2. Extraer Utilidad de la API para un Ejecutivo Comercial
  window.getApiUtilidadForEjecutivo = function (ejName) {
    if (!window.API_UTILIDAD_DATA || !window.API_UTILIDAD_DATA.length) return null;
    const norm = String(ejName || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    const tokens = norm.split(/\s+/).filter(t => t.length > 2);

    const found = window.API_UTILIDAD_DATA.find(r => {
      const desc = String(r.Descripcion || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
      const matches = tokens.filter(tok => desc.includes(tok));
      return matches.length >= 2 || (tokens.length === 1 && matches.length === 1) || desc.includes(norm) || norm.includes(desc);
    });

    if (!found) return null;

    return {
      utilidad: Number(found.Valor_Utilidad) || 0,
      mercancia: Number(found.Valor_Mercancia) || 0,
      costo: Number(found.Valor_Costo) || 0,
      porcentaje: Number(found.Porcentaje_Utilidad) || 0
    };
  };

  // 3. Extraer Utilidad de la API para un Director (Suma del equipo)
  window.getApiUtilidadForDirector = function (dirName, execsList) {
    if (!window.API_UTILIDAD_DATA || !window.API_UTILIDAD_DATA.length) return null;
    let totalUtilidad = 0;
    let totalMercancia = 0;
    let matchedCount = 0;

    (execsList || []).forEach(exec => {
      const apiData = window.getApiUtilidadForEjecutivo(exec);
      if (apiData) {
        totalUtilidad += apiData.utilidad;
        totalMercancia += apiData.mercancia;
        matchedCount++;
      }
    });

    return matchedCount > 0 ? {
      utilidad: totalUtilidad,
      mercancia: totalMercancia,
      matchedCount: matchedCount
    } : null;
  };

  // 4. Iniciar carga al cargar el módulo
  loadApiUtilidadCache();

})();
