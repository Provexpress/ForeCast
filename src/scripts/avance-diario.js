// ============================================================================
// MÓDULO DE AVANCE DIARIO DE CUOTA Y UTILIDAD COMERCIAL (FORECAST 2026)
// Consume la API de Utilidad Comercial con soporte dinámico para cualquier mes
// ============================================================================

(function () {
  window.API_UTILIDAD_DATA = [];
  window.API_UTILIDAD_MESES = {};

  
  
  // Función para realizar consulta EN VIVO a través del Proxy local sin CORS (Idéntico a Postman)
  async function fetchLiveUtilidadFromApi(fechaInicial, fechaFinal) {
    try {
      // 1. Intentar Proxy local (http://localhost:5001/api/utilidad)
      const proxyUrl = `http://localhost:5001/api/utilidad?fechaInicial=${fechaInicial}&fechaFinal=${fechaFinal}`;
      const proxyResp = await fetch(proxyUrl, { method: "GET" });
      if (proxyResp.ok) {
        const proxyData = await proxyResp.json();
        if (proxyData && proxyData.response && proxyData.response.length) {
          console.log("⚡ [PROXY EN VIVO] Petición en tiempo real exitosa:", proxyData.response.length, "vendedores");
          return proxyData.response;
        }
      }
    } catch (e) {
      console.log("ℹ️ Proxy local no activo o no alcanzable, intentando petición directa.");
    }

    try {
      // 2. Intentar petición directa por si el navegador lo permite
      const authResp = await fetch("http://152.200.146.226:50010/api/getKey", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username: "powerbi", password: "3xpress#2025" })
      });
      
      if (!authResp.ok) return null;
      const authData = await authResp.json();
      const token = authData.token || authData.access_token;
      if (!token) return null;

      const dataResp = await fetch("http://152.200.146.226:50010/consultas/api/consultaUtilidadComercialesDashboardPBI", {
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

      if (!dataResp.ok) return null;
      const resultData = await dataResp.json();
      return resultData.response || [];
    } catch (e) {
      return null;
    }
  }

  // 1. Cargar datos de la API de Utilidad en el cliente (Intenta EN VIVO primero, luego cache)
  async function loadApiUtilidadCache() {
    try {
      const now = new Date();
      const year = now.getFullYear();
      const month = String(now.getMonth() + 1).padStart(2, "0");
      const day = String(now.getDate()).padStart(2, "0");
      const curMonthKey = `${year}-${month}`;
      const fechaInicial = `${curMonthKey}-01`;
      const fechaFinal = `${year}-${month}-${day}`;

      // Intentar peticion directa EN VIVO a la API (tipo Postman)
      const liveRows = await fetchLiveUtilidadFromApi(fechaInicial, fechaFinal);
      if (liveRows && liveRows.length) {
        window.API_UTILIDAD_DATA = liveRows;
        if (!window.API_UTILIDAD_MESES) window.API_UTILIDAD_MESES = {};
        window.API_UTILIDAD_MESES[curMonthKey] = {
          fechaInicial,
          fechaFinal,
          totalVendedores: liveRows.length,
          vendedores: liveRows
        };
        console.log('⚡ Conexión EN VIVO a la API de Power BI establecida exitosamente:', liveRows.length, 'registros');
      }

      // Cargar cache de respaldo
      const cachePath = 'src/data/api-utilidad-cache.json?v=' + Date.now();
      const resp = await fetch(cachePath);
      if (resp.ok) {
        const data = await resp.json();
        if (!window.API_UTILIDAD_DATA || !window.API_UTILIDAD_DATA.length) {
          window.API_UTILIDAD_DATA = data.vendedores || [];
        }
        window.API_UTILIDAD_MESES = Object.assign({}, data.meses || {}, window.API_UTILIDAD_MESES || {});
        console.log('✅ Base de datos de API cargada en ForeCast con', Object.keys(window.API_UTILIDAD_MESES).length, 'meses');
      }

      if (typeof window.refreshAvanceDiarioViews === 'function') {
        window.refreshAvanceDiarioViews();
      }
    } catch (e) {
      console.warn('⚠️ Error al cargar datos de API Utilidad:', e.message);
    }
  }


  const API_NAME_ALIASES = {
    "johanna mojica": "Jasbleidy Johana Mojica",
    "jasbleidy mojica": "Jasbleidy Johana Mojica",
    "jasbleidy johana mojica": "Jasbleidy Johana Mojica",
    "johana mojica": "Jasbleidy Johana Mojica",
    
    "yovanny herrera": "Jair Yovanny Herrea",
    "jair yovanny herrera": "Jair Yovanny Herrea",
    "jair yovanny herrea": "Jair Yovanny Herrea",
    
    "rosmira rojas": "Rosmira Rojas Puentes",
    "mario reyes": "Mario Reyes Gutierrez",
    "wilson sanchez": "Wilson Fernando Sanchez Monroy",
    "maria eugenia cruz": "Maria Eugenia Cruz Herrera",
    "javier cortes": "Javier Antonio Cortes Murcia",
    "rosa mendoza": "Rosa Maria Mendoza Mendoza",
    "mariela ramirez": "Mariela Ramírez Castro",
    "jenny gonzalez": "Jenny Alexandra Gonzalez Buitrago",
    "julieth galindo": "Julieth Milena Galindo Fino",
    "angela torres": "Angela Rocio Torres Matallana",
    "yurany andrea vargas": "Yurany Andrea Vargas Soler",
    "andrea vargas": "Yurany Andrea Vargas Soler",
    "alejandra velasquez": "Maria Alejandra Velásquez Espinosa",
    "fernando quinonez": "Fernando Alberto Quiñonez",
    "fernando alberto quinonez": "Fernando Alberto Quiñonez",
    "johanna jaime": "Johanna Jaime Murcia",
    "johanna jaime murcia": "Johanna Jaime Murcia",
    "dayana chala": "Dayana Marcela Chala Rodríguez",
    "cesar cespedes": "Cesar Augusto Cespedes Sabroso",
    "daniel galindo": "Daniel Galindo Giron",
    "gina garcia": "Gina Paola Garcia Quito",
    "karent carrillo": "Karent Carrillo Marin",
    "lington linares": "Lington Linares Linares",
    "angelica alvarez": "Maria Angelica Alvarez Morales",
    "andres pena": "Freddy Andres Peña Sanchez",
    "freddy pena": "Freddy Andres Peña Sanchez",
    "tatiana parra": "Angie Tatiana Parra Durán",
    "claudia triana": "Claudia Patricia Triana Olaya",
    "dilma cuesta": "Dilma Constanza Cuesta Rubiano",
    "juan martinez": "Juan David Martínez Pedraza",
    "astrid jimenez": "Leidy Astrid Jimenez Ossa",
    "maria paola briceno": "Maria Paola Briceño Muñoz",
    "paola briceno": "Maria Paola Briceño Muñoz",
    "dafne ruiz": "Dafne Lizeth Ruiz Bernal",
    "jessica valencia": "Jessica Lorena Valencia Isaza",
    "lorena valencia": "Jessica Lorena Valencia Isaza",
    "jhonatan acevedo": "Jhonatan Steven Acevedo Fonseca",
    "steven acevedo": "Jhonatan Steven Acevedo Fonseca",
    "camilo hernandez": "Jhonatan Camilo Hernandez Martinez",
    "yeison urrego": "Yeison Alonso Urrego Cortes",
    "diana castro": "Diana Catalina Castro Castro"
  };

  function normalizeName(str) {
    return String(str || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  }

  // 2. Extraer Utilidad de la API para un Ejecutivo Comercial según el mes seleccionado
  window.getApiUtilidadForEjecutivo = function (ejName, selectedMonthKey) {
    let list = [];
    
    if (selectedMonthKey) {
      if (window.API_UTILIDAD_MESES && window.API_UTILIDAD_MESES[selectedMonthKey]) {
        list = window.API_UTILIDAD_MESES[selectedMonthKey].vendedores || [];
      } else {
        return null;
      }
    } else if (Array.isArray(window.API_UTILIDAD_DATA) && window.API_UTILIDAD_DATA.length) {
      list = window.API_UTILIDAD_DATA;
    }

    if (!list.length) return null;

    const normInput = normalizeName(ejName);
    const aliasTarget = API_NAME_ALIASES[normInput] ? normalizeName(API_NAME_ALIASES[normInput]) : null;

    // 1. Buscar por Alias explícito
    let found = aliasTarget ? list.find(r => normalizeName(r.Descripcion) === aliasTarget) : null;

    // 2. Buscar por coincidencia exacta de nombre limpio
    if (!found) {
      found = list.find(r => normalizeName(r.Descripcion) === normInput);
    }

    // 3. Buscar por puntuación de tokens
    if (!found) {
      const tokens = normInput.split(/\s+/).filter(t => t.length > 2);
      let bestMatch = null;
      let maxScore = 0;

      list.forEach(r => {
        const normApi = normalizeName(r.Descripcion);
        const apiTokens = normApi.split(/\s+/).filter(t => t.length > 2);
        
        let score = 0;
        tokens.forEach(t => {
          if (apiTokens.includes(t)) score += 2;
          else if (normApi.includes(t)) score += 1;
        });

        if (normApi.includes(normInput) || normInput.includes(normApi)) score += 3;

        if (score > maxScore) {
          maxScore = score;
          bestMatch = r;
        }
      });

      if (maxScore >= 2) found = bestMatch;
    }

    if (!found) return null;

    return {
      utilidad: Number(found.Valor_Utilidad) || 0,
      mercancia: Number(found.Valor_Mercancia) || 0,
      costo: Number(found.Valor_Costo) || 0,
      porcentaje: Number(found.Porcentaje_Utilidad) || 0
    };
  };

  // 3. Extraer Utilidad de la API para un Director según el mes seleccionado
  window.getApiUtilidadForDirector = function (dirName, execsList, selectedMonthKey) {
    let totalUtilidad = 0;
    let totalMercancia = 0;
    let matchedCount = 0;

    (execsList || []).forEach(exec => {
      const apiData = window.getApiUtilidadForEjecutivo(exec, selectedMonthKey);
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


  // 5. Exportar Informe Profesional Excel para Director
  
  // Helper para resolver los ejecutivos de cualquier director
  window.getDirectorExecs = function (dirName) {
    const norm = String(dirName || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    
    // 1. ESTRUCTURA_COMERCIAL_2026
    const est = window.ESTRUCTURA_COMERCIAL_2026;
    if (est && est.directores && est.ejecutivos) {
      let grupoId = null;
      for (const info of Object.values(est.directores)) {
        const dNorm = String(info.nombre || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
        if (dNorm.includes(norm) || norm.includes(dNorm)) {
          grupoId = info.grupo;
          break;
        }
      }
      if (grupoId) {
        return Object.values(est.ejecutivos)
          .filter(e => e.grupo === grupoId)
          .map(e => e.nombre);
      }
    }

    // 2. Fallback ALL_DATA & LOADED_FILES_BY_DIR
    const allData = window.ALL_DATA || [];
    const execsWithData = allData
      .filter(r => {
        const d = String(r['DIRECTOR'] || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
        return d.includes(norm) || norm.includes(d);
      })
      .map(r => (r['COMERCIAL'] || '').trim())
      .filter(Boolean);

    const loadedFiles = window.LOADED_FILES_BY_DIR || {};
    let execsFromFiles = [];
    for (const [dKey, files] of Object.entries(loadedFiles)) {
      const dNorm = String(dKey || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
      if (dNorm.includes(norm) || norm.includes(dNorm)) {
        execsFromFiles = (files || []).map(f => f.name.replace(/\.(xlsx|xls)$/i, '').trim()).filter(Boolean);
      }
    }

    return [...new Set([...execsWithData, ...execsFromFiles])].sort();
  };

  window.exportDirectorReportToExcel = async function () {
    if (typeof ExcelJS === 'undefined') {
      alert('Cargando librería de Excel, por favor intenta de nuevo en un momento.');
      return;
    }

    const dirSelect = document.getElementById('sel-director');
    const dirName = dirSelect ? dirSelect.value : 'Director';
    const mesSelect = document.getElementById('sel-dir-mes');
    const mesKey = mesSelect ? mesSelect.value : '2026-07';
    const mesText = mesSelect && mesSelect.options[mesSelect.selectedIndex] ? mesSelect.options[mesSelect.selectedIndex].text : mesKey;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Avance ' + dirName.substring(0, 15));

    sheet.views = [{ showGridLines: true }];

    // Encabezado Corporativo
    sheet.mergeCells('A1:G2');
    const titleCell = sheet.getCell('A1');
    titleCell.value = `PROVEXPRESS S.A.S. — INFORME DE AVANCE DE CUOTA Y UTILIDAD (API POWER BI)\nDirector: ${dirName} | Periodo: ${mesText}`;
    titleCell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A2B6B' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

    sheet.addRow([]); // Fila vacia

    // Encabezados de Tabla
    const headerRow = sheet.addRow([
      'Ejecutivo Comercial',
      'Cuota Mensual ($)',
      'Utilidad Lograda API ($)',
      'Ventas Totales ($)',
      'Faltante Meta ($)',
      '% Cumplimiento',
      'Estado'
    ]);

    headerRow.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.height = 26;

    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F7891' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD4DDED' } },
        bottom: { style: 'medium', color: { argb: 'FF1A2B6B' } },
        left: { style: 'thin', color: { argb: 'FFD4DDED' } },
        right: { style: 'thin', color: { argb: 'FFD4DDED' } }
      };
    });

    const execs = (typeof window.getDirectorExecs === 'function') ? window.getDirectorExecs(dirName) : [];
    
    let totalCuota = 0;
    let totalUtilidad = 0;
    let totalVentas = 0;

    execs.forEach((e) => {
      const eCuota = (typeof getExecutiveCuota === 'function') ? getExecutiveCuota(e) : 18000000;
      const eApiData = typeof window.getApiUtilidadForEjecutivo === 'function' ? window.getApiUtilidadForEjecutivo(e, mesKey) : null;
      const eUtilidad = eApiData ? eApiData.utilidad : 0;
      const eVentas = eApiData ? eApiData.mercancia : 0;
      const ePct = eCuota > 0 ? (eUtilidad / eCuota) : 0;
      const eFaltante = Math.max(0, eCuota - eUtilidad);

      totalCuota += eCuota;
      totalUtilidad += eUtilidad;
      totalVentas += eVentas;

      let estadoLabel = 'En Riesgo';
      let estadoBg = 'FFFEE2E2';
      let estadoFg = 'FF991B1B';

      if (ePct >= 1.0) {
        estadoLabel = 'Meta Alcanzada';
        estadoBg = 'FFD1FAE5';
        estadoFg = 'FF065F46';
      } else if (ePct >= 0.5) {
        estadoLabel = 'En Camino';
        estadoBg = 'FFFEF3C7';
        estadoFg = 'FF92400E';
      }

      const row = sheet.addRow([
        e,
        eCuota,
        eUtilidad,
        eVentas,
        eFaltante,
        ePct,
        estadoLabel
      ]);

      row.height = 22;

      row.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
      row.getCell(1).font = { name: 'Arial', size: 10, bold: true };

      row.getCell(2).numberFormat = '$ #,##0';
      row.getCell(2).alignment = { vertical: 'middle', horizontal: 'right' };

      row.getCell(3).numberFormat = '$ #,##0';
      row.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' };
      row.getCell(3).font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF047857' } };

      row.getCell(4).numberFormat = '$ #,##0';
      row.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' };

      row.getCell(5).numberFormat = '$ #,##0';
      row.getCell(5).alignment = { vertical: 'middle', horizontal: 'right' };

      row.getCell(6).numberFormat = '0.0%';
      row.getCell(6).alignment = { vertical: 'middle', horizontal: 'right' };
      row.getCell(6).font = { name: 'Arial', size: 10, bold: true };

      const statusCell = row.getCell(7);
      statusCell.alignment = { vertical: 'middle', horizontal: 'center' };
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: estadoBg } };
      statusCell.font = { name: 'Arial', size: 9, bold: true, color: { argb: estadoFg } };

      row.eachCell((c) => {
        c.border = {
          top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        };
      });
    });

    // Fila de Totales
    const totalPct = totalCuota > 0 ? (totalUtilidad / totalCuota) : 0;
    const totalFaltante = Math.max(0, totalCuota - totalUtilidad);

    let totalEstado = 'En Riesgo';
    let totalBg = 'FFFEE2E2';
    let totalFg = 'FF991B1B';
    if (totalPct >= 1.0) { totalEstado = 'Meta Alcanzada'; totalBg = 'FFD1FAE5'; totalFg = 'FF065F46'; }
    else if (totalPct >= 0.5) { totalEstado = 'En Camino'; totalBg = 'FFFEF3C7'; totalFg = 'FF92400E'; }

    const totalRow = sheet.addRow([
      'TOTAL GRUPO',
      totalCuota,
      totalUtilidad,
      totalVentas,
      totalFaltante,
      totalPct,
      totalEstado
    ]);

    totalRow.height = 25;
    totalRow.eachCell((c, colNum) => {
      c.font = { name: 'Arial', size: 10, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
      c.border = {
        top: { style: 'medium', color: { argb: 'FF1A2B6B' } },
        bottom: { style: 'double', color: { argb: 'FF1A2B6B' } },
        left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
      };
      if (colNum >= 2 && colNum <= 5) {
        c.numberFormat = '$ #,##0';
        c.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (colNum === 6) {
        c.numberFormat = '0.0%';
        c.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (colNum === 7) {
        c.alignment = { vertical: 'middle', horizontal: 'center' };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: totalBg } };
        c.font = { name: 'Arial', size: 9, bold: true, color: { argb: totalFg } };
      }
    });

    sheet.columns = [
      { width: 32 },
      { width: 22 },
      { width: 24 },
      { width: 22 },
      { width: 22 },
      { width: 18 },
      { width: 18 }
    ];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Reporte_Avance_Director_${dirName.replace(/\s+/g, '_')}_${mesKey}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  // 6. Exportar Informe Profesional Excel para Ejecutivo Individual
  window.exportEjecutivoReportToExcel = async function () {
    if (typeof ExcelJS === 'undefined') {
      alert('Cargando librería de Excel, por favor intenta de nuevo en un momento.');
      return;
    }

    const ejSelect = document.getElementById('sel-ejecutivo');
    const ejName = ejSelect ? ejSelect.value : 'Ejecutivo';
    const mesSelect = document.getElementById('sel-ej-mes');
    const mesKey = mesSelect ? mesSelect.value : '2026-07';
    const mesText = mesSelect && mesSelect.options[mesSelect.selectedIndex] ? mesSelect.options[mesSelect.selectedIndex].text : mesKey;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Cuota ' + ejName.substring(0, 15));

    sheet.views = [{ showGridLines: true }];

    sheet.mergeCells('A1:E2');
    const titleCell = sheet.getCell('A1');
    titleCell.value = `PROVEXPRESS S.A.S. — INFORME DE AVANCE INDIVIDUAL (API POWER BI)\nEjecutivo: ${ejName} | Periodo: ${mesText}`;
    titleCell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A2B6B' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

    sheet.addRow([]);

    const eCuota = (typeof getExecutiveCuota === 'function') ? getExecutiveCuota(ejName) : 18000000;
    const eApiData = typeof window.getApiUtilidadForEjecutivo === 'function' ? window.getApiUtilidadForEjecutivo(ejName, mesKey) : null;
    const eUtilidad = eApiData ? eApiData.utilidad : 0;
    const eVentas = eApiData ? eApiData.mercancia : 0;
    const ePct = eCuota > 0 ? (eUtilidad / eCuota) : 0;
    const eFaltante = Math.max(0, eCuota - eUtilidad);

    let estadoLabel = 'En Riesgo';
    let estadoBg = 'FFFEE2E2';
    let estadoFg = 'FF991B1B';
    if (ePct >= 1.0) { estadoLabel = 'Meta Alcanzada'; estadoBg = 'FFD1FAE5'; estadoFg = 'FF065F46'; }
    else if (ePct >= 0.5) { estadoLabel = 'En Camino'; estadoBg = 'FFFEF3C7'; estadoFg = 'FF92400E'; }

    const headerRow = sheet.addRow(['Métrica Comercial', 'Valor / Detalle']);
    headerRow.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.eachCell(c => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F7891' } };
      c.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const rows = [
      ['Cuota Mensual asignada', eCuota, '$ #,##0'],
      ['Utilidad Lograda API (Avance)', eUtilidad, '$ #,##0'],
      ['Ventas Totales ($ Facturado)', eVentas, '$ #,##0'],
      ['Faltante para la Meta', eFaltante, '$ #,##0'],
      ['Porcentaje de Cumplimiento', ePct, '0.0%'],
      ['Estado de Avance', estadoLabel, '@']
    ];

    rows.forEach(r => {
      const addedRow = sheet.addRow([r[0], r[1]]);
      addedRow.height = 22;
      addedRow.getCell(1).font = { name: 'Arial', size: 10, bold: true };
      
      const valCell = addedRow.getCell(2);
      valCell.alignment = { vertical: 'middle', horizontal: 'right' };
      if (r[2] === '$ #,##0') {
        valCell.numberFormat = '$ #,##0';
        valCell.font = { name: 'Arial', size: 10, bold: true };
      } else if (r[2] === '0.0%') {
        valCell.numberFormat = '0.0%';
        valCell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF047857' } };
      } else if (r[0] === 'Estado de Avance') {
        valCell.alignment = { vertical: 'middle', horizontal: 'center' };
        valCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: estadoBg } };
        valCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: estadoFg } };
      }
    });

    sheet.columns = [
      { width: 32 },
      { width: 28 }
    ];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Reporte_Cuota_Ejecutivo_${ejName.replace(/\s+/g, '_')}_${mesKey}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

})();