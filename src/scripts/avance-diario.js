// ============================================================================
// MÓDULO DE AVANCE DIARIO DE CUOTA Y UTILIDAD COMERCIAL (FORECAST 2026)
// Consume la API de Utilidad Comercial con soporte dinámico para cualquier mes
// ============================================================================

(function () {
  window.API_UTILIDAD_DATA = [];
  window.API_UTILIDAD_MESES = {};

  // 1. Cargar datos de la API de Utilidad en el cliente
  async function loadApiUtilidadCache() {
    try {
      const cachePath = 'src/data/api-utilidad-cache.json?v=' + Date.now();
      const resp = await fetch(cachePath);
      if (resp.ok) {
        const data = await resp.json();
        window.API_UTILIDAD_DATA = data.vendedores || [];
        window.API_UTILIDAD_MESES = data.meses || {};
        console.log('✅ API Utilidad Comercial cargada en ForeCast con', Object.keys(window.API_UTILIDAD_MESES).length, 'meses sincronizados');
        
        if (typeof window.refreshAvanceDiarioViews === 'function') {
          window.refreshAvanceDiarioViews();
        }
      }
    } catch (e) {
      console.warn('⚠️ No se pudo cargar api-utilidad-cache.json, usando fallback local.', e.message);
    }
  }

  // 2. Extraer Utilidad de la API para un Ejecutivo Comercial según el mes seleccionado
  window.getApiUtilidadForEjecutivo = function (ejName, selectedMonthKey) {
    let list = [];
    
    if (selectedMonthKey && window.API_UTILIDAD_MESES && window.API_UTILIDAD_MESES[selectedMonthKey]) {
      list = window.API_UTILIDAD_MESES[selectedMonthKey].vendedores || [];
    } else if (Array.isArray(window.API_UTILIDAD_DATA) && window.API_UTILIDAD_DATA.length) {
      list = window.API_UTILIDAD_DATA;
    }

    if (!list.length) return null;

    const norm = String(ejName || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    const tokens = norm.split(/\s+/).filter(t => t.length > 2);

    const found = list.find(r => {
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