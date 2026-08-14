const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const main = fs.readFileSync(path.join(root, 'src', 'scripts', 'main.js'), 'utf8');
const fondos = fs.readFileSync(path.join(root, 'src', 'scripts', 'fondos-marketing.js'), 'utf8');

assert.match(html, /id="tab-fondos"[\s\S]*onclick="showPage\('fondos',this\)"/);
assert.match(html, /src\/styles\/main\.css\?v=20260814-fondos-chart1/);
assert.match(html, /<div class="page" id="page-fondos"><\/div>/);
assert.match(html, /id="program-channel-workbook-table"><\/div>[\s\S]*<\/section>[\s\S]*<\/section>[\s\S]*<\/div>[\s\S]*PAGE: FONDOS MARKETING[\s\S]*id="page-fondos"/);
assert.match(html, /src\/scripts\/fondos-marketing\.js\?v=20260814-fondos-chart1/);
assert.match(html, /src\/scripts\/main\.js\?v=20260814-fondos-chart1/);

assert.match(main, /if\(page === 'fondos'\)[\s\S]*renderFondosMarketing\(\)/);
assert.match(main, /id === 'fondos'/);
assert.match(main, /const canViewFondos = role === 'gerencia'/);
assert.match(main, /function loadFondosMarketingModule\(\)/);
assert.match(main, /function showFondosMarketingError\(error\)/);
assert.match(main, /window\.renderFondosMarketing = renderFondosMarketing/);

assert.match(fondos, /function init\(\)[\s\S]*renderLayout\(container\)[\s\S]*updateDashboardView\(\)/);
assert.match(fondos, /return CURRENT_USER\.role === 'gerencia'/);
assert.match(fondos, /<section class="fondos-header" aria-labelledby="fondos-page-title">/);
assert.doesNotMatch(fondos, /<header class="fondos-header">/);
assert.doesNotMatch(fondos, /let _rendered = false/);
assert.match(fondos, /El tablero de Fondos Marketing no quedó insertado/);
assert.match(fondos, /sourceFileName: 'FONDOS DE MERCADEO NINI_final\.xlsx'/);
assert.match(fondos, /cop: 62627700|D11/);
assert.match(fondos, /excludeFromConsolidated: true/);
assert.match(fondos, /function parseFinalWorkbook\(workbook, fileName, lastModified\)/);
assert.match(fondos, /async function loadFromSharePoint\(siteId, token\)/);
assert.match(fondos, /fondosBarValueLabels/);
assert.match(fondos, /fondosDoughnutValueLabels/);
assert.match(fondos, /fondos-bar-card/);
assert.match(fondos, /badgeText = hasIncome[\s\S]*'Sin ingreso'/);
assert.match(fondos, /maxBarThickness: 34/);
assert.match(fondos, /'Champion'/);
assert.match(fondos, /generateLabels\(chart\)[\s\S]*doughnutLabels\.map/);

console.log('Fondos Marketing: navegación, permisos, caché e inicialización validados.');
