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

function getFormattedMonthRange() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return {
    fechaInicial: `${year}-${month}-01`,
    fechaFinal: `${year}-${month}-${day}`
  };
}

async function syncUtilidadData() {
  console.log("📥 Conectando con la API de Utilidad Comercial en tiempo real...");
  const range = getFormattedMonthRange();
  
  try {
    const auth = await postJSON(API_AUTH_URL, API_CREDS);
    const token = auth.token || auth.access_token;
    if (!token) throw new Error("No se obtuvo token de autenticación de la API");

    const dataResp = await postJSON(API_DATA_URL, {
      Fecha_Inicial: range.fechaInicial,
      Fecha_Final: range.fechaFinal,
      Tipo_Utilidad: "venta"
    }, token);

    const rows = dataResp.response || [];
    console.log(`✅ API de Utilidad respondió exitosamente: ${rows.length} vendedores procesados`);

    const outputDir = path.join(__dirname, "../src/data");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    const payload = {
      updatedAt: new Date().toISOString(),
      fechaInicial: range.fechaInicial,
      fechaFinal: range.fechaFinal,
      totalVendedores: rows.length,
      vendedores: rows
    };

    const filePath = path.join(outputDir, "api-utilidad-cache.json");
    fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), "utf8");
    console.log(`💾 Cache sincronizado guardado en: ${filePath}`);
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
