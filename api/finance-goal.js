import { getGoal, setGoal } from './goals-store.js';

export default async function handler(req, res) {
  // Configuración de CORS segura basada en orígenes autorizados
  const origin = req.headers.origin || req.headers.referer || "";
  const allowedOrigins = [
    "https://tableros-area-financiera.vercel.app",
    "https://provexpress.sharepoint.com",
    "http://localhost",
    "http://127.0.0.1"
  ];
  
  const isAllowed = allowedOrigins.some(o => origin.startsWith(o)) || 
                    /https?:\/\/localhost(:\d+)?/.test(origin) || 
                    /https?:\/\/127\.0\.0\.1(:\d+)?/.test(origin) || 
                    origin.endsWith(".vercel.app");
                    
  if (isAllowed) {
    res.setHeader("Access-Control-Allow-Origin", origin.startsWith("http") ? origin.replace(/\/$/, "") : "*");
  } else {
    res.setHeader("Access-Control-Allow-Origin", "https://tableros-area-financiera.vercel.app");
  }

  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  
  if (req.method === "OPTIONS") return res.status(204).end();

  // Procesamiento de peticiones
  if (req.method === "GET") {
    const period = req.query.period;
    if (!period || !/^\d{4}-\d{2}$/.test(period)) {
      return res.status(400).json({ ok: false, error: "Formato de periodo inválido (requerido YYYY-MM)" });
    }
    try {
      const goal = await getGoal(period);
      return res.json({ ok: true, period, goal });
    } catch (e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  if (req.method === "POST") {
    const body = req.body || {};
    const { period, goal } = body;
    if (!period || !/^\d{4}-\d{2}$/.test(period)) {
      return res.status(400).json({ ok: false, error: "Formato de periodo inválido (requerido YYYY-MM)" });
    }
    try {
      await setGoal(period, goal);
      return res.json({ ok: true, period, goal: parseFloat(goal) || 0 });
    } catch (e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  return res.status(405).json({ ok: false, error: "Método no permitido" });
}
