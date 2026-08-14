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
      set fillStyle(value) {}
    };
    FakeChart.instances.push(this);
    (config.plugins || []).forEach(plugin => plugin.afterDatasetsDraw?.(this));
  }

  getDatasetMeta(datasetIndex) {
    const values = this.data.datasets[datasetIndex]?.data || [];
    return {
      data: values.map((value, index) => ({
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
assert.match(elements.get('fondos-kpis-container').innerHTML, /194\.115\.588/);
assert.match(elements.get('fondos-summary-tbody').innerHTML, /TOTAL CONSOLIDADO/);
assert.match(elements.get('fondos-brand-detail-container').innerHTML, /HP/);
assert.equal(FakeChart.instances.length, 2);
assert.equal(FakeChart.instances[0].config.plugins[0].id, 'fondosBarValueLabels');
assert.equal(FakeChart.instances[1].config.plugins[0].id, 'fondosDoughnutValueLabels');
assert.match(FakeChart.instances[0].options.plugins.tooltip.callbacks.label({ datasetIndex: 1, dataIndex: 0, dataset: { label: 'Ejecutado (COP)' }, raw: 58269668 }), /93,0%/);
assert.match(FakeChart.instances[1].options.plugins.tooltip.callbacks.label({ label: 'Eventos', raw: 31321068 }), /14,8%/);
const doughnutLegend = FakeChart.instances[1].options.plugins.legend.labels.generateLabels(FakeChart.instances[1]);
assert.equal(doughnutLegend.length, 4);
assert.match(doughnutLegend[3].text, /Champion · \$28,3M · 13,4%/);

for(const id of [...elements.keys()]) {
  if(id !== 'page-fondos') elements.delete(id);
}
elements.get('page-fondos').innerHTML = '';
assert.doesNotThrow(() => context.FondosMarketingModule.init());
assert.match(elements.get('page-fondos').innerHTML, /Fondos de Mercadeo \(MDF\)/);
assert.match(elements.get('fondos-kpis-container').innerHTML, /Ingresos Totales/);

console.log('Fondos Marketing: renderizado inicial y reconstrucción completa validados.');
