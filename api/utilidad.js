const DEFAULT_PBI_BASE_URL = "http://152.200.146.226:50010";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=300");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const baseUrl = process.env.PBI_API_URL || DEFAULT_PBI_BASE_URL;
  const fechaInicial = req.query?.fechaInicial || req.body?.fechaInicial || "2026-08-01";
  const fechaFinal = req.query?.fechaFinal || req.body?.fechaFinal || "2026-08-03";

  try {
    // 1. Auth Key
    const authResp = await fetch(`${baseUrl}/api/getKey`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username: "powerbi", password: "3xpress#2025" })
    });

    if (!authResp.ok) {
      const errText = await authResp.text().catch(() => "");
      return res.status(authResp.status).json({ ok: false, error: `Auth falló (${authResp.status}): ${errText}` });
    }

    const authData = await authResp.json();
    const token = authData.token || authData.access_token;

    if (!token) {
      return res.status(500).json({ ok: false, error: "API Power BI no retornó token" });
    }

    // 2. Fetch Utilidad
    const dataResp = await fetch(`${baseUrl}/consultas/api/consultaUtilidadComercialesDashboardPBI`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({
        Fecha_Inicial: fechaInicial,
        Fecha_Final: fechaFinal,
        Tipo_Utilidad: "venta"
      })
    });

    if (!dataResp.ok) {
      const errText = await dataResp.text().catch(() => "");
      return res.status(dataResp.status).json({ ok: false, error: `Consulta falló (${dataResp.status}): ${errText}` });
    }

    const resultData = await dataResp.json();
    return res.json({
      ok: true,
      fechaInicial,
      fechaFinal,
      totalVendedores: (resultData.response || []).length,
      response: resultData.response || []
    });

  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
}
