const http = require("http");

const PORT = 5001;
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

const server = http.createServer(async (req, res) => {
  // Permitir CORS a cualquier origen (Navegador)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    res.writeHead(200);
    res.end();
    return;
  }

  const reqUrl = new URL(req.url, `http://${req.headers.host}`);

  if (reqUrl.pathname === "/api/utilidad") {
    try {
      const now = new Date();
      const year = now.getFullYear();
      const month = String(now.getMonth() + 1).padStart(2, "0");
      const day = String(now.getDate()).padStart(2, "0");
      
      const fechaInicial = reqUrl.searchParams.get("fechaInicial") || `${year}-${month}-01`;
      const fechaFinal = reqUrl.searchParams.get("fechaFinal") || `${year}-${month}-${day}`;

      console.log(`📡 [PROXY EN VIVO] Peticion del navegador recibida. Consultando API (${fechaInicial} al ${fechaFinal})...`);

      const auth = await postJSON(API_AUTH_URL, API_CREDS);
      const token = auth.token || auth.access_token;
      if (!token) throw new Error("No token from Power BI API");

      const dataResp = await postJSON(API_DATA_URL, {
        Fecha_Inicial: fechaInicial,
        Fecha_Final: fechaFinal,
        Tipo_Utilidad: "venta"
      }, token);

      const rows = dataResp.response || [];
      console.log(`✅ [PROXY EN VIVO] API respondio con ${rows.length} registros para el navegador.`);

      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ status: "success", count: rows.length, response: rows }));
    } catch (error) {
      console.error("❌ [PROXY EN VIVO] Error consultando API:", error.message);
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ status: "error", message: error.message }));
    }
  } else {
    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("404 Not Found");
  }
});

server.listen(PORT, () => {
  console.log(`============================================================================`);
  console.log(`⚡ PROVEXPRESS FORECAST - PROXY SERVER EN VIVO ACTIVO EN http://localhost:${PORT}`);
  console.log(`============================================================================`);
  console.log(`Servicio listo para intermediar peticiones en vivo del navegador sin bloqueo CORS.`);
});
