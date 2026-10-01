/**
 * sync-api-utilidad.js
 * Actualiza src/data/api-utilidad-cache.json con datos de la API Power BI.
 * 
 * Estrategia defensiva:
 *  - Si un mes falla, conserva los datos previos del cache para ese mes
 *  - Si la API está completamente caída, NO sobreescribe el archivo
 *  - Solo actualiza meses que tengan datos válidos (al menos 1 vendedor)
 *  - Timeout de 20s por request para evitar colgar el proceso
 */

const http  = require("http");
const fs    = require("fs");
const path  = require("path");

const API_AUTH_URL = "http://152.200.146.226:50010/api/getKey";
const API_DATA_URL = "http://152.200.146.226:50010/consultas/api/consultaUtilidadComercialesDashboardPBI";
const API_CREDS    = { username: "powerbi", password: "3xpress#2025" };
const TIMEOUT_MS   = 20000; // 20 segundos por request

const CACHE_FILE   = path.join(__dirname, "../src/data/api-utilidad-cache.json");

// ── HTTP util ──────────────────────────────────────────────────────────────

function postJSON(url, data, token) {
  return new Promise((resolve, reject) => {
    const u       = new URL(url);
    const bodyStr = JSON.stringify(data);
    const options = {
      hostname: u.hostname,
      port:     u.port,
      path:     u.pathname,
      method:   "POST",
      headers:  { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(bodyStr) },
      timeout:  TIMEOUT_MS
    };
    if (token) options.headers["Authorization"] = "Bearer " + token;

    const req = http.request(options, (res) => {
      let raw = "";
      res.on("data", chunk => raw += chunk);
      res.on("end", () => {
        try { resolve(JSON.parse(raw)); }
        catch (e) { reject(new Error("JSON parse error: " + raw.slice(0, 120))); }
      });
    });

    req.on("timeout", () => { req.destroy(); reject(new Error("Request timeout (" + TIMEOUT_MS + "ms)")); });
    req.on("error",   reject);
    req.write(bodyStr);
    req.end();
  });
}

// ── Rangos de meses ────────────────────────────────────────────────────────

function getMonthRanges() {
  const now             = new Date();
  const currentYear     = now.getFullYear();
  const currentMonthNum = now.getMonth() + 1;
  const currentDay      = String(now.getDate()).padStart(2, "0");
  const ranges          = {};

  // Mes actual (del 1 al día de hoy)
  const curMonthKey = `${currentYear}-${String(currentMonthNum).padStart(2, "0")}`;
  ranges[curMonthKey] = {
    fechaInicial: `${curMonthKey}-01`,
    fechaFinal:   `${curMonthKey}-${currentDay}`
  };

  // Meses anteriores del año (completos)
  for (let m = 1; m < currentMonthNum; m++) {
    const mStr         = String(m).padStart(2, "0");
    const monthKey     = `${currentYear}-${mStr}`;
    const lastDay      = new Date(currentYear, m, 0).getDate();
    ranges[monthKey] = {
      fechaInicial: `${monthKey}-01`,
      fechaFinal:   `${monthKey}-${String(lastDay).padStart(2, "0")}`
    };
  }

  return ranges;
}

// ── Sync principal ─────────────────────────────────────────────────────────

async function syncUtilidadData() {
  console.log("📥 Conectando con la API de Utilidad Comercial...");
  const monthRanges = getMonthRanges();

  // Cargar cache existente como respaldo
  let cacheAnterior = { meses: {} };
  if (fs.existsSync(CACHE_FILE)) {
    try {
      cacheAnterior = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8"));
      console.log("  💾 Cache anterior cargado (" + Object.keys(cacheAnterior.meses || {}).join(", ") + ")");
    } catch (e) {
      console.warn("  ⚠️  No se pudo leer cache anterior:", e.message);
    }
  }

  // Obtener token
  let token;
  try {
    const auth = await postJSON(API_AUTH_URL, API_CREDS);
    token = auth.token || auth.access_token;
    if (!token) throw new Error("Respuesta sin token: " + JSON.stringify(auth).slice(0, 120));
    console.log("  🔑 Token obtenido OK");
  } catch (e) {
    console.error("❌ No se pudo autenticar con la API:", e.message);
    console.error("   El cache existente NO fue modificado.");
    process.exit(1);
  }

  // Consultar cada mes — defensivo: si falla, conserva el mes anterior del cache
  const mesesData     = { ...(cacheAnterior.meses || {}) };
  let   exitosamente  = 0;
  let   conFallback   = 0;

  for (const [monthKey, range] of Object.entries(monthRanges)) {
    process.stdout.write(`  └─ ${monthKey} (${range.fechaInicial} → ${range.fechaFinal})... `);
    try {
      const dataResp = await postJSON(API_DATA_URL, {
        Fecha_Inicial: range.fechaInicial,
        Fecha_Final:   range.fechaFinal,
        Tipo_Utilidad: "venta"
      }, token);

      const rows = dataResp.response || [];

      if (rows.length === 0 && monthKey !== Object.keys(monthRanges)[0]) {
        // Mes histórico sin filas — puede ser error de la API, conservar previo
        const hayPrevio = mesesData[monthKey]?.vendedores?.length > 0;
        console.log(hayPrevio ? `⚠️  0 vendedores — conservando datos previos` : "⚠️  0 vendedores");
        if (hayPrevio) { conFallback++; continue; }
      }

      mesesData[monthKey] = {
        fechaInicial:   range.fechaInicial,
        fechaFinal:     range.fechaFinal,
        totalVendedores: rows.length,
        vendedores:     rows
      };

      const utilTotal = rows.reduce((s, v) => s + (v.Valor_Utilidad || 0), 0);
      console.log(`✅ ${rows.length} vendedores | $${Math.round(utilTotal / 1e6)}M`);
      exitosamente++;

    } catch (e) {
      const hayPrevio = mesesData[monthKey]?.vendedores?.length > 0;
      console.log(`❌ ${e.message}${hayPrevio ? " — conservando datos previos" : ""}`);
      if (hayPrevio) conFallback++;
    }
  }

  // No guardar si NO se actualizó absolutamente nada
  if (exitosamente === 0) {
    console.error("\n❌ Ningún mes se actualizó correctamente. El cache NO fue modificado.");
    process.exit(1);
  }

  // Guardar
  const currentMonthKey = Object.keys(monthRanges)[0];
  const payload = {
    updatedAt:        new Date().toISOString(),
    currentMonthKey:  currentMonthKey,
    vendedores:       mesesData[currentMonthKey]?.vendedores || [],
    meses:            mesesData
  };

  const outputDir = path.dirname(CACHE_FILE);
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
  fs.writeFileSync(CACHE_FILE, JSON.stringify(payload, null, 2), "utf8");

  console.log(`\n✅ Cache guardado — ${exitosamente} meses actualizados, ${conFallback} conservados del cache anterior.`);
  console.log(`   Archivo: ${CACHE_FILE}`);
  return payload;
}

if (require.main === module) {
  syncUtilidadData().catch(e => { console.error(e); process.exit(1); });
}

module.exports = { syncUtilidadData };
