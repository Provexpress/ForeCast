const http = require("http");
const fs = require("fs");
const path = require("path");

const API_AUTH_URL = "http://152.200.146.226:50010/api/getKey";
const API_DATA_URL = "http://152.200.146.226:50010/consultas/api/consultaUtilidadComercialesDashboardPBI";
const API_CREDS = { username: "powerbi", password: "3xpress#2025" };

function postJSON(url, data, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const bodyStr = JSON.stringify(data);
    const options = {
      hostname: u.hostname,
      port: u.port,
      path: u.pathname,
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(bodyStr)
      }
    };
    if (token) options.headers["Authorization"] = "Bearer " + token;
    const req = http.request(options, (res) => {
      let responseData = "";
      res.on("data", chunk => responseData += chunk);
      res.on("end", () => {
        try {
          resolve(JSON.parse(responseData));
        } catch (e) {
          reject(e);
        }
      });
    });
    req.on("error", reject);
    req.write(bodyStr);
    req.end();
  });
}

function getMonthRanges() {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonthNum = now.getMonth() + 1;
  const currentDay = String(now.getDate()).padStart(2, "0");

  const ranges = {};

  // 1. Mes actual (Julio 2026: del 1 al dia de hoy)
  const curMonthKey = `${currentYear}-${String(currentMonthNum).padStart(2, "0")}`;
  ranges[curMonthKey] = {
    fechaInicial: `${curMonthKey}-01`,
    fechaFinal: `${curMonthKey}-${currentDay}`
  };

  // 2. Meses anteriores del año (Junio, Mayo, Abril, Marzo, Febrero, Enero)
  for (let m = 1; m < currentMonthNum; m++) {
    const mStr = String(m).padStart(2, "0");
    const monthKey = `${currentYear}-${mStr}`;
    const lastDayOfMonth = new Date(currentYear, m, 0).getDate();
    ranges[monthKey] = {
      fechaInicial: `${monthKey}-01`,
      fechaFinal: `${monthKey}-${String(lastDayOfMonth).padStart(2, "0")}`
    };
  }

  return ranges;
}

async function syncUtilidadData() {
  console.log("📥 Conectando con la API de Utilidad Comercial para todos los meses del año...");
  const monthRanges = getMonthRanges();
  
  try {
    const auth = await postJSON(API_AUTH_URL, API_CREDS);
    const token = auth.token || auth.access_token;
    if (!token) throw new Error("No se obtuvo token de autenticación de la API");

    const mesesData = {};

    for (const [monthKey, range] of Object.entries(monthRanges)) {
      console.log(`  └─ Consultando periodo ${monthKey} (${range.fechaInicial} al ${range.fechaFinal})...`);
      const dataResp = await postJSON(API_DATA_URL, {
        Fecha_Inicial: range.fechaInicial,
        Fecha_Final: range.fechaFinal,
        Tipo_Utilidad: "venta"
      }, token);

      const rows = dataResp.response || [];
      mesesData[monthKey] = {
        fechaInicial: range.fechaInicial,
        fechaFinal: range.fechaFinal,
        totalVendedores: rows.length,
        vendedores: rows
      };
    }

    const outputDir = path.join(__dirname, "../src/data");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    // Determinar vendedores del mes actual para compatibilidad
    const currentMonthKey = Object.keys(monthRanges)[0];
    const payload = {
      updatedAt: new Date().toISOString(),
      currentMonthKey: currentMonthKey,
      vendedores: mesesData[currentMonthKey]?.vendedores || [],
      meses: mesesData
    };

    const filePath = path.join(outputDir, "api-utilidad-cache.json");
    fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), "utf8");
    console.log(`✅ Sincronización multi-mes completada exitosamente. Guardado en: ${filePath}`);
    return payload;
  } catch (error) {
    console.error("❌ Error al sincronizar con la API de Utilidad:", error.message);
    throw error;
  }
}

if (require.main === module) {
  syncUtilidadData();
}

module.exports = { syncUtilidadData };
