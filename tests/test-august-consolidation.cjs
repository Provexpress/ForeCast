const fs = require('fs');
const cache = require('../src/data/api-utilidad-cache.json');

// Mock window
global.window = global;

// Load estructura-comercial-2026.js
const estCode = fs.readFileSync('./src/config/estructura-comercial-2026.js', 'utf8');
eval(estCode);

// Load avance-diario.js
const avanceCode = fs.readFileSync('./src/scripts/avance-diario.js', 'utf8');
eval(avanceCode);

// Set cache in window.API_UTILIDAD_MESES
window.API_UTILIDAD_DATA = cache.vendedores || [];
window.API_UTILIDAD_MESES = cache.meses || {};

const dirs = [
  { name: 'Rafael Novoa', key: 'Grupo Rafael Novoa' },
  { name: 'Angélica Caballero', key: 'Grupo Maria Angelica caballero' },
  { name: 'Óscar Beltrán', key: 'Grupo Oscar Beltran' },
  { name: 'Miller Romero', key: 'Gupo Miller Romero' }
];

console.log('Testing August (2026-08) matching for all 38 executives:');
let grandUtilidad = 0;
let grandVentas = 0;
let totalMatched = 0;
let totalExecs = 0;

dirs.forEach(dirObj => {
  const execs = window.getDirectorExecs(dirObj.name);
  console.log(`\n=== Director: ${dirObj.name} (${execs.length} execs) ===`);
  let groupUtil = 0;
  let groupVentas = 0;
  execs.forEach(e => {
    totalExecs++;
    const data = window.getApiUtilidadForEjecutivo(e, '2026-08');
    if (data) {
      totalMatched++;
      groupUtil += data.utilidad;
      groupVentas += data.mercancia;
      console.log(`  ✓ ${e.padEnd(28)} | Util: $${data.utilidad.toLocaleString('es-CO').padStart(14)} | Ventas: $${data.mercancia.toLocaleString('es-CO').padStart(16)}`);
    } else {
      console.log(`  ✗ ${e.padEnd(28)} | NO MATCH / 0 DATA`);
    }
  });
  console.log(`Subtotal ${dirObj.name}: Utilidad = $${groupUtil.toLocaleString('es-CO')}, Ventas = $${groupVentas.toLocaleString('es-CO')}`);
  grandUtilidad += groupUtil;
  grandVentas += groupVentas;
});

console.log('\n=============================================');
console.log(`TOTAL CONSOLIDADO 2026-08:`);
console.log(`Ejecutivos evaluados: ${totalExecs}`);
console.log(`Ejecutivos con match en API: ${totalMatched}/${totalExecs}`);
console.log(`Utilidad Total: $${grandUtilidad.toLocaleString('es-CO')}`);
console.log(`Ventas Totales: $${grandVentas.toLocaleString('es-CO')}`);
