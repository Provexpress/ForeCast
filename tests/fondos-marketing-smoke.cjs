const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const main = fs.readFileSync(path.join(root, 'src', 'scripts', 'main.js'), 'utf8');
const fondos = fs.readFileSync(path.join(root, 'src', 'scripts', 'fondos-marketing.js'), 'utf8');

assert.match(html, /id="tab-fondos"[\s\S]*onclick="showPage\('fondos',this\)"/);
assert.match(html, /<div class="page" id="page-fondos"><\/div>/);
assert.match(html, /src\/scripts\/fondos-marketing\.js\?v=20260813-fondos3/);
assert.match(html, /src\/scripts\/main\.js\?v=20260813-fondos4/);

assert.match(main, /if\(page === 'fondos'\)[\s\S]*FondosMarketingModule\.init\(\)/);
assert.match(main, /id === 'fondos'/);
assert.match(main, /const canViewFondos = role === 'gerencia'/);

assert.match(fondos, /function init\(\)[\s\S]*renderLayout\(container\)[\s\S]*updateDashboardView\(\)/);
assert.match(fondos, /return CURRENT_USER\.role === 'gerencia'/);

console.log('Fondos Marketing: navegación, permisos, caché e inicialización validados.');
