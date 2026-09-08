const fs = require('fs');
const cache = require('../src/data/api-utilidad-cache.json');

global.window = global;

const estCode = fs.readFileSync('./src/config/estructura-comercial-2026.js', 'utf8');
eval(estCode);

const avanceCode = fs.readFileSync('./src/scripts/avance-diario.js', 'utf8');
eval(avanceCode);

window.API_UTILIDAD_DATA = cache.vendedores || [];
window.API_UTILIDAD_MESES = cache.meses || {};

const months = Object.keys(window.API_UTILIDAD_MESES).sort();
const dirs = ['Rafael Novoa', 'Angélica Caballero', 'Óscar Beltrán', 'Miller Romero'];

console.log('Testing all months for all executives:');
months.forEach(m => {
  let monthUtil = 0;
  let monthVentas = 0;
  let matched = 0;
  let total = 0;

  dirs.forEach(d => {
    const execs = window.getDirectorExecs(d);
    execs.forEach(e => {
      total++;
      const data = window.getApiUtilidadForEjecutivo(e, m);
      if (data) {
        matched++;
        monthUtil += data.utilidad;
        monthVentas += data.mercancia;
      }
    });
  });

  console.log(`Month ${m}: ${matched}/${total} execs matched | Total Utilidad: $${monthUtil.toLocaleString('es-CO')} | Total Ventas: $${monthVentas.toLocaleString('es-CO')}`);
});
