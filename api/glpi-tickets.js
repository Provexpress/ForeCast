const DEFAULT_GLPI_BASE_URL = "http://152.200.168.131:50010/glpi/apirest.php";
const DEFAULT_APP_TOKEN = "JmXIjj4ngfJ2NLM81Z36h8JCQLF6aUVh6bS1pC2f";
const DEFAULT_USER_TOKEN = "t6Fjn3XoEJB6FKLFo1KUdY2VIUZnmDZKhcjuxmB1";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, Session-Token, App-Token");
  res.setHeader("Cache-Control", "s-maxage=300, stale-while-revalidate=600");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const baseUrl = process.env.GLPI_API_URL || DEFAULT_GLPI_BASE_URL;
  const appToken = process.env.GLPI_APP_TOKEN || DEFAULT_APP_TOKEN;
  const userToken = process.env.GLPI_USER_TOKEN || DEFAULT_USER_TOKEN;

  let sessionToken = null;
  try {
    // 1. initSession
    const initRes = await fetch(`${baseUrl}/initSession`, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `user_token ${userToken}`,
        "App-Token": appToken
      }
    });

    if (!initRes.ok) {
      const errText = await initRes.text().catch(() => "");
      return res.status(initRes.status).json({ ok: false, error: `initSession falló (${initRes.status}): ${errText}` });
    }

    const initData = await initRes.json();
    sessionToken = initData.session_token;
    if (!sessionToken) {
      return res.status(500).json({ ok: false, error: "GLPI no entregó un session_token válido" });
    }

    // 2. Fetch user map for names
    let userMap = {};
    try {
      const usersRes = await fetch(`${baseUrl}/User?range=0-1000`, {
        headers: {
          "Content-Type": "application/json",
          "Session-Token": sessionToken,
          "App-Token": appToken
        }
      });
      if (usersRes.ok) {
        const userList = await usersRes.json();
        (userList || []).forEach(u => {
          if (!u || u.id === undefined) return;
          const fullName = [u.firstname, u.realname].filter(Boolean).join(" ").trim() || u.name || "";
          if (fullName) {
            userMap[u.id] = fullName;
            userMap[String(u.id)] = fullName;
          }
        });
      }
    } catch (uErr) {
      console.warn("[GLPI Proxy] No se pudo descargar mapa de usuarios:", uErr.message);
    }

    // 3. Fetch tickets paginated
    const FIELDS = { id: 2, titulo: 1, estado: 12, fecha_apertura: 15, fecha_solucion: 17, fecha_cierre: 16, categoria: 7, solicitante: 4, tecnico: 5, grupo_tecnico: 8 };
    const PAGE_SIZE_FETCH = 500;
    let allRawTickets = [];
    let start = 0;
    let totalCount = null;

    do {
      const params = new URLSearchParams();
      params.set("criteria[0][field]", FIELDS.fecha_apertura);
      params.set("criteria[0][searchtype]", "morethan");
      params.set("criteria[0][value]", "2026-01-01");
      params.set("expand_dropdowns", "true");
      params.set("range", `${start}-${start + PAGE_SIZE_FETCH - 1}`);

      Object.values(FIELDS).forEach((fieldId, idx) => {
        params.set(`forcedisplay[${idx}]`, fieldId);
      });

      const searchUrl = `${baseUrl}/search/Ticket?${params.toString()}`;
      const searchRes = await fetch(searchUrl, {
        headers: {
          "Content-Type": "application/json",
          "Session-Token": sessionToken,
          "App-Token": appToken
        }
      });

      if (!searchRes.ok) {
        break;
      }

      const json = await searchRes.json();
      totalCount = typeof json.totalcount === "number" ? json.totalcount : null;
      const batch = json.data || [];
      allRawTickets = allRawTickets.concat(batch);
      start += PAGE_SIZE_FETCH;
    } while (totalCount !== null && start < totalCount);

    return res.json({
      ok: true,
      totalCount: totalCount || allRawTickets.length,
      rawTickets: allRawTickets,
      userMap,
      updatedAt: new Date().toISOString()
    });

  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  } finally {
    if (sessionToken) {
      try {
        await fetch(`${baseUrl}/killSession`, {
          headers: {
            "Session-Token": sessionToken,
            "App-Token": appToken
          }
        });
      } catch (kErr) {}
    }
  }
}
