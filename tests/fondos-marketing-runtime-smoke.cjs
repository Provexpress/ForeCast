const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const elements = new Map();

class FakeElement {
  constructor(id = '') {
    this.id = id;
    this._innerHTML = '';
    this.style = {};
  }

  set innerHTML(value) {
    this._innerHTML = String(value);
    for(const match of this._innerHTML.matchAll(/id="([^"]+)"/g)) {
      if(!elements.has(match[1])) elements.set(match[1], new FakeElement(match[1]));
    }
  }

  get innerHTML() {
    return this._innerHTML;
  }

  addEventListener() {}
  getAttribute() { return null; }
  querySelector(selector) {
    if(selector === '.fondos-app-wrapper' && this._innerHTML.includes('class="fondos-app-wrapper"')) return {};
    return null;
  }
  querySelectorAll() { return []; }
}

elements.set('page-fondos', new FakeElement('page-fondos'));

const document = {
  getElementById(id) {
    return elements.get(id) || null;
  }
};

class FakeChart {
  static instances = [];
  static defaults = {
    plugins: {
      legend: {
        labels: {
          generateLabels(chart) {
            return (chart.data.labels || []).map((text, index) => ({ text, index }));
          }
        }
      }
    }
  };

  constructor(canvas, config) {
    this.canvas = canvas;
    this.config = config;
    this.data = config.data;
    this.options = config.options;
    this.chartArea = { left: 0, right: 320, top: 0, bottom: 260 };
    this.width = 900;
    this.ctx = {
      save() {},
      restore() {},
      strokeText() {},
      fillText() {},
      set textAlign(value) {},
      set textBaseline(value) {},
      set lineJoin(value) {},
      set strokeStyle(value) {},
      set lineWidth(value) {},
      set font(value) {},
      set fillStyle(value) {},
      beginPath() {},
      roundRect() {},
      fill() {},
      stroke() {},
      measureText(value) { return { width: String(value).length * 5 }; }
    };
    FakeChart.instances.push(this);
    (config.plugins || []).forEach(plugin => plugin.afterDatasetsDraw?.(this));
  }

  getDatasetMeta(datasetIndex) {
    const values = this.data.datasets[datasetIndex]?.data || [];
    return {
      data: values.map((value, index) => ({
        y: value < 0 ? 210 : 80,
        base: 180,
        tooltipPosition() {
          return { x: 40 + index * 38 + datasetIndex * 8, y: value < 0 ? 200 : 90 };
        }
      }))
    };
  }

  getDataVisibility() {
    return true;
  }

  destroy() {
    this.destroyed = true;
  }
}

const context = {
  console,
  document,
  CURRENT_USER: { role: 'gerencia', email: 'especialista.preventa@provexpress.com.co' },
  Intl,
  JSON,
  Math,
  Number,
  Date,
  Chart: FakeChart,
  getComputedStyle() {
    return { getPropertyValue() { return '#172033'; } };
  },
  setTimeout,
  clearTimeout
};
context.window = context;

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'src', 'scripts', 'fondos-marketing.js'), 'utf8');
vm.runInNewContext(source, context, { filename: 'fondos-marketing.js' });

assert.ok(context.FondosMarketingModule, 'El módulo debe registrarse en window');
assert.equal(context.FondosMarketingModule.canAccess(), true);
assert.doesNotThrow(() => context.FondosMarketingModule.init());
assert.match(elements.get('page-fondos').innerHTML, /Fondos de Mercadeo \(MDF\)/);
assert.match(elements.get('fondos-kpis-container').innerHTML, /Ingresos Totales/);
assert.match(elements.get('fondos-kpis-container').innerHTML, /218\.615\.588/);
assert.doesNotMatch(elements.get('fondos-kpis-container').innerHTML, /Comisión Dir\. Mercadeo/);
assert.match(elements.get('fondos-rio-funding').innerHTML, /Valor por cupo/);
assert.match(elements.get('fondos-rio-funding').innerHTML, /6\.500\.000/);
assert.match(elements.get('fondos-rio-funding').innerHTML, /HP/);
assert.match(elements.get('fondos-rio-funding').innerHTML, /26\.000\.000/);
assert.match(elements.get('fondos-rio-funding').innerHTML, /No ha ingresado dinero/);
assert.match(elements.get('fondos-rio-funding').innerHTML, /130\.000\.000/);
assert.match(elements.get('fondos-summary-tbody').innerHTML, /TOTAL CONSOLIDADO/);
assert.match(elements.get('fondos-brand-detail-container').innerHTML, /HP/);
assert.equal(FakeChart.instances.length, 2);
assert.equal(FakeChart.instances[0].config.plugins[0].id, 'fondosBarValueLabels');
assert.equal(FakeChart.instances[1].config.plugins[0].id, 'fondosDoughnutValueLabels');
assert.equal(FakeChart.instances[0].options.scales.y.max, 90000000);
assert.equal(FakeChart.instances[0].options.scales.y.min, -40000000);
assert.equal(FakeChart.instances[0].data.datasets[0].maxBarThickness, 34);
assert.match(FakeChart.instances[0].options.plugins.tooltip.callbacks.label({ datasetIndex: 1, dataIndex: 0, dataset: { label: 'Ejecutado' }, raw: 51769668 }), /51\.769\.668/);
assert.match(FakeChart.instances[0].options.plugins.tooltip.callbacks.afterBody([{ dataIndex: 0 }]), /74,4%/);
assert.match(FakeChart.instances[1].options.plugins.tooltip.callbacks.label({ label: 'Eventos', raw: 31321068 }), /14,3%/);
const doughnutLegend = FakeChart.instances[1].options.plugins.legend.labels.generateLabels(FakeChart.instances[1]);
assert.equal(doughnutLegend.length, 6);
assert.match(doughnutLegend[3].text, /Champion · \$28,0M · 12,8%/);
assert.match(doughnutLegend[4].text, /POP · \$20,9M · 9,5%/);

function makeSheet(cells) {
  return Object.fromEntries(Object.entries(cells).map(([address, value]) => [address, { v: value, w: typeof value === 'string' ? value : undefined }]));
}

const finalWorkbook = {
  SheetNames: ['hp', 'Dell', 'Lenovo', 'Intel', 'Microsoft', 'Cisco', 'Champion', 'Total Fondos '],
  Sheets: {
    hp: makeSheet({
      B2: 2200, C2: 3861, D2: 8494200, E2: '26Q1', F2: 108962,
      B3: 500, C3: 4019, D3: 2009500, E3: '26Q1', F3: 110438,
      B4: 1500, C4: 3861, D4: 5791500, E4: '26Q1', F4: 106826,
      B5: 1500, C5: 3861, D5: 5791500, E5: '26Q1', F5: 107237,
      B6: 3000, C6: 3668, D6: 11004000, E6: '26Q2', F6: 112747,
      B7: 3000, C7: 3668, D7: 11004000, E7: '26Q2', F7: 113461,
      B8: 1000, C8: 3861, D8: 3861000, E8: '26Q2', F8: 106957,
      B9: 3000, C9: 3668, D9: 11004000, E9: '26Q3', F9: 113701,
      B10: 1000, C10: 3668, D10: 3668000, E10: '26Q3', F10: 112763,
      B11: 2000, C11: 3500, D11: 7000000, E11: '26Q4', F11: 'Pendiente factura',
      D12: 69627700, D13: 17858032,
      H2: 2710400, I2: 'HP suministros', H3: 486000, I3: 'Liga Z', H4: 730085, I4: 'Workstation', H5: 5943183, I5: 'Nueva era',
      J2: 2500000, K2: 'Concurso Supplies', J3: 3000000, K3: 'Concurso portafolio', J4: 6000000, K4: 'Campaña', J5: 1500000, K5: 'Camisetas', J6: 300000, K6: 'Bono Adidas', J7: 2600000, K7: 'Nintendo',
      L2: 26000000, M2: '1 cupo cruzado con rebate', H12: 51769668
    }),
    Dell: makeSheet({
      C2: '2026', D2: 'Q3', E2: 'Agosto', F2: 1414.49, G2: 3513.54, H2: 4969868,
      C3: '2026', D3: 'Q4', E3: 'Noviembre', F3: 1000, G3: 3810.99, H3: 3810990,
      C4: '2027', D4: 'Q1', E4: 'Enero', F4: 2864, G4: 3066.99, H4: 8783850,
      C5: '2027', D5: 'Q2', E5: 'Mayo', F5: 1811, G5: 3962.55, H5: 7176180,
      C8: 'Proposal', F8: 6500, G8: 3500, H8: 22750000,
      C9: 'Proposal pendiente', F9: 3000, G9: 3500, H9: 10500000,
      H11: 57990888, H12: 11063438,
      L2: 8874800, M2: 'Infraestructura', L3: 6452650, M3: 'Cómputo', O2: 3000000, P2: 'Televisor', O3: 2600000, P3: 'Play Station', Q2: 26000000, S6: 46927450
    }),
    Lenovo: makeSheet({
      A2: 3349, B2: 3500, C2: 11721500, D2: 'Validar NC', A4: 1906, B4: 3500, C4: 6671000, D4: 'Ingreso', C5: 18392500, C6: 1233550,
      H2: 4408950, I2: 'Infraestructura', J2: 2500000, K2: 'Legion', J3: 500000, K3: 'Bono', L2: 9750000, M2: 17158950
    }),
    Intel: makeSheet({
      B2: 'H1', C2: 4990, D2: 3500, E2: 17465000, B3: 'H2', C3: 5000, D3: 3500, E3: 17500000,
      G2: 1715000, H2: 'Evento', I2: 6000000, J2: 'Capacitación', K2: 9750000
    }),
    Microsoft: makeSheet({
      D2: 'LOL', E2: 3000, F2: 3500, G2: 10500000, D3: 'INGRAM', E3: 4897, F3: 3500, G3: 17139500,
      N2: 17139500, O2: 'Cruce de cuentas Licencias', P2: 9750000
    }),
    Cisco: makeSheet({ B2: 3000, C2: 3333.3333333333335, D2: 10000000, G2: 9750000, H1: 'Saldo cruza negocio', H2: 250000 }),
    Champion: makeSheet({ D2: 'Q1', E2: 2000, F2: 7000000, G2: 2000, H2: 7000000, D3: 'Q2', E3: 2000, F3: 7000000, G3: 2000, H3: 7000000 }),
    'Total Fondos ': makeSheet({
      H3: 'POP', I3: 20891156, A9: 'Total', B9: 218615588, C9: 170210568, D9: 48405020,
      A15: 6500000, B15: 20, C15: 130000000,
      A17: 'Recaudo Rio', B17: 'Valor', C17: 'Observacion',
      A18: 'Hp', B18: 26000000,
      A19: 'Dell', B19: 26000000,
      A20: 'Lenovo', B20: 9750000,
      A21: 'Intel', B21: 9750000,
      A22: 'Cisco', B22: 9750000,
      A23: 'Microsoft', B23: 9750000, C23: 'No ha ingresado Dinero',
      A24: 'Provexpress', B24: 19500000, C24: 'Asumio Provexpress 3 Cupos',
      A25: 'Provexpress mercadeo', B25: 19500000, C25: 'Conserguir Dinero Mercadeo',
      A28: 'Total Fabricas', B28: 91000000,
      A29: 'Asumio Provexpress 3 Cupos', B29: 19500000,
      A30: 'Conserguir  Dinero Mercadeo', B30: 19500000,
      A31: 'Total', B31: 130000000
    })
  }
};

const parsedFinal = context.FondosMarketingModule.parseFinalWorkbook(finalWorkbook, 'FONDOS DE MERCADEO NINI_final.xlsx', '2026-08-18T12:55:47Z');
assert.equal(parsedFinal.lastUpdate, '2026-08-18');
assert.equal(parsedFinal.brands.find(brand => brand.id === 'hp').incomes.length, 10);
assert.equal(parsedFinal.brands.find(brand => brand.id === 'intel').incomes.length, 2);
assert.equal(parsedFinal.brands.find(brand => brand.id === 'microsoft').outflows.reduce((sum, item) => sum + item.cop, 0), 26889500);
assert.equal(parsedFinal.brands.find(brand => brand.id === 'pop').outflows[0].cop, 20891156);
assert.equal(parsedFinal.rioFunding.unitCost, 6500000);
assert.equal(parsedFinal.rioFunding.totalSeats, 20);
assert.equal(parsedFinal.rioFunding.factories, 91000000);
assert.equal(parsedFinal.rioFunding.provexpress, 19500000);
assert.equal(parsedFinal.rioFunding.marketingPending, 19500000);
assert.equal(parsedFinal.rioFunding.total, 130000000);
assert.equal(parsedFinal.rioFunding.contributions.length, 8);
assert.equal(parsedFinal.rioFunding.contributions[0].name, 'Hp');
assert.equal(parsedFinal.rioFunding.contributions[0].seats, 4);
assert.equal(parsedFinal.rioFunding.contributions[5].observation, 'No ha ingresado Dinero');

for(const id of [...elements.keys()]) {
  if(id !== 'page-fondos') elements.delete(id);
}
elements.get('page-fondos').innerHTML = '';
assert.doesNotThrow(() => context.FondosMarketingModule.init());
assert.match(elements.get('page-fondos').innerHTML, /Fondos de Mercadeo \(MDF\)/);
assert.match(elements.get('fondos-kpis-container').innerHTML, /Ingresos Totales/);

console.log('Fondos Marketing: renderizado inicial y reconstrucción completa validados.');
