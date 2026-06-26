const DEFAULT_CACHE_URL = "https://tableros-area-financiera.vercel.app/_cache/ventas-pbi.json";
const DEFAULT_PBI_BASE_URL = "http://152.200.146.226:50010";

const PBI_ENDPOINTS = {
  ventas: "/consultas/api/consultaVentasDashboardPBI",
};

const FINANCE_CATEGORY_MARGINS = {
  "Servicios": 0.3435,
  "Renta de Equipos": 0.2412,
  PAC: 0.2349,
  Suministros: 0.1558,
  "Tecnología": 0.1206,
  Licenciamiento: 0.1065,
};
const FINANCE_FALLBACK_MARGIN_CATEGORY = "Servicios";

function getEnv(name, fallback = "") {
  return process.env[name] || fallback;
}

function joinUrl(baseUrl, endpoint) {
  return `${String(baseUrl || "").replace(/\/+$/, "")}/${String(endpoint || "").replace(/^\/+/, "")}`;
}

function getBogotaDateParts(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Bogota",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  return Object.fromEntries(parts.filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
}

function getTodayBogotaIso() {
  const parts = getBogotaDateParts();
  return `${parts.year}-${parts.month}-${parts.day}`;
}

function getMonthStart(period) {
  return `${period}-01`;
}

function getQueryValue(req, name) {
  const value = req.query?.[name];
  return Array.isArray(value) ? value[0] : value;
}

function isIsoDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""));
}

function normalizePeriod(value) {
  const candidate = String(value || "").slice(0, 7);
  if (/^\d{4}-\d{2}$/.test(candidate)) return candidate;
  return getTodayBogotaIso().slice(0, 7);
}

function normalizeDateRange(req) {
  const period = normalizePeriod(getQueryValue(req, "period"));
  const startDate = isIsoDate(getQueryValue(req, "startDate")) ? getQueryValue(req, "startDate") : getMonthStart(period);
  const today = getTodayBogotaIso();
  const requestedEnd = getQueryValue(req, "endDate");
  const endDate = isIsoDate(requestedEnd) ? requestedEnd : period === today.slice(0, 7) ? today : `${period}-31`;
  return {
    period,
    startDate,
    endDate,
  };
}

function parseFlexibleNumber(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "number") return Number.isFinite(value) ? value : fallback;
  const text = String(value)
    .trim()
    .replace(/\s/g, "")
    .replace(/\$/g, "");
  if (!text) return fallback;
  const normalized = text.includes(",") && text.includes(".")
    ? text.replace(/\./g, "").replace(",", ".")
    : text.replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeText(value) {
  return cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function normalizeCategory(value) {
  const normalized = normalizeText(value);
  if (!normalized) return "Ventas";
  if (normalized.includes("renta")) return "Renta de Equipos";
  if (normalized.includes("servicio")) return "Servicios";
  if (normalized.includes("licenciamiento")) return "Licenciamiento";
  if (normalized.includes("suministro")) return "Suministros";
  if (normalized.includes("tecnolog")) return "Tecnología";
  if (normalized === "pac" || normalized.includes("pac")) return "PAC";
  return cleanText(value) || "Ventas";
}

function getFinanceMarginConfig(category) {
  const hasExplicitMargin = Object.prototype.hasOwnProperty.call(FINANCE_CATEGORY_MARGINS, category);
  const marginSource = hasExplicitMargin ? category : FINANCE_FALLBACK_MARGIN_CATEGORY;
  return {
    marginPct: FINANCE_CATEGORY_MARGINS[marginSource] || 0,
    marginSource,
    marginIsFallback: !hasExplicitMargin,
  };
}

function parseSpreadsheetDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && Number.isFinite(value)) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    epoch.setUTCDate(epoch.getUTCDate() + Math.floor(value));
    return epoch.toISOString().slice(0, 10);
  }
  const raw = cleanText(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
  const dmy = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (dmy) {
    const year = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3];
    return `${year}-${String(dmy[2]).padStart(2, "0")}-${String(dmy[1]).padStart(2, "0")}`;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10);
}

function readJsonResponse(response) {
  return response.text().then((text) => {
    if (!text) return null;
    try {
      return JSON.parse(text);
    } catch {
      return text;
    }
  });
}

function extractRows(payload) {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.result)) return payload.result;
  if (Array.isArray(payload?.resultado)) return payload.resultado;
  if (Array.isArray(payload?.rows)) return payload.rows;
  if (Array.isArray(payload?.payload?.data)) return payload.payload.data;
  if (Array.isArray(payload?.response)) return payload.response;
  if (Array.isArray(payload?.response?.data)) return payload.response.data;
  if (Array.isArray(payload?.response?.result)) return payload.response.result;
  if (Array.isArray(payload?.response?.resultado)) return payload.response.resultado;
  if (Array.isArray(payload?.response?.rows)) return payload.response.rows;
  if (typeof payload?.response === "string") {
    try {
      return extractRows(JSON.parse(payload.response));
    } catch {
      return [];
    }
  }
  return [];
}

function unwrapCacheEnvelope(candidate) {
  if (candidate?.cacheVersion && candidate?.payload) {
    return {
      rows: candidate.payload.data || [],
      meta: candidate.payload.meta || {},
      generatedAt: candidate.generatedAt || candidate.payload.meta?.cacheGeneratedAt || null,
      source: "cache",
    };
  }
  return {
    rows: candidate?.data || [],
    meta: candidate?.meta || {},
    generatedAt: candidate?.generatedAt || candidate?.meta?.cacheGeneratedAt || null,
    source: "cache",
  };
}

async function requestPbiToken(baseUrl) {
  const user = getEnv("PBI_API_USER");
  const password = getEnv("PBI_API_PASSWORD");
  const userField = getEnv("PBI_API_USER_FIELD", "username");
  const passwordField = getEnv("PBI_API_PASSWORD_FIELD", "password");
  if (!user || !password) throw new Error("Faltan PBI_API_USER y/o PBI_API_PASSWORD.");

  const response = await fetch(joinUrl(baseUrl, "/api/getKey"), {
    method: "POST",
    headers: { Accept: "application/json", "Content-Type": "application/json" },
    body: JSON.stringify({ [userField]: user, [passwordField]: password }),
  });
  const payload = await readJsonResponse(response);
  if (!response.ok) throw new Error(`Token PBI HTTP ${response.status}`);
  const token =
    payload?.token ||
    payload?.access_token ||
    payload?.key ||
    payload?.apiKey ||
    payload?.data?.token ||
    payload?.data?.access_token ||
    payload?.data?.key ||
    "";
  if (!token) throw new Error("No se pudo identificar el token PBI.");
  return token;
}

function buildAuthHeaders(token) {
  const tokenHeader = getEnv("PBI_API_TOKEN_HEADER", "Authorization");
  if (tokenHeader.toLowerCase() === "x-api-key") return { "x-api-key": token };
  return { [tokenHeader]: `Bearer ${token}` };
}

async function fetchPbiRows() {
  const baseUrl = getEnv("PBI_API_BASE_URL", DEFAULT_PBI_BASE_URL);
  const token = await requestPbiToken(baseUrl);
  const response = await fetch(joinUrl(baseUrl, PBI_ENDPOINTS.ventas), {
    method: "GET",
    headers: {
      Accept: "application/json",
      ...buildAuthHeaders(token),
    },
  });
  const payload = await readJsonResponse(response);
  if (!response.ok) throw new Error(`Ventas PBI HTTP ${response.status}`);
  return {
    rows: extractRows(payload),
    meta: { sourceName: "API de ventas PBI" },
    generatedAt: new Date().toISOString(),
    source: "api",
  };
}

async function loadRows() {
  const cacheUrl = getEnv("FINANCE_VENTAS_CACHE_URL", DEFAULT_CACHE_URL);
  if (cacheUrl) {
    try {
      const response = await fetch(cacheUrl, { headers: { Accept: "application/json" } });
      if (response.ok) {
        const payload = await response.json();
        const unwrapped = unwrapCacheEnvelope(payload);
        if (unwrapped.rows.length) return unwrapped;
      }
    } catch (error) {
      console.warn("[finance-summary] cache no disponible", error.message);
    }
  }
  return fetchPbiRows();
}

function normalizeCachedDocument(row) {
  const sign = Number(row.signoDocumento || 1) < 0 ? -1 : 1;
  const totalOriginal = Math.abs(parseFlexibleNumber(row.totalOriginal ?? row.total, 0));
  const total = parseFlexibleNumber(row.total, totalOriginal * sign);
  const cost = sign > 0 ? parseFlexibleNumber(row.costoProducto, 0) : 0;
  return {
    fechaIso: String(row.fechaIso || "").slice(0, 10),
    sign,
    total,
    totalOriginal,
    cost,
    category: normalizeCategory(row.categoria),
    document: row.numeroDocumento || row.folio || "",
    client: row.cliente || row.proveedor || "",
  };
}

function normalizeRawPbiRows(rows) {
  const buckets = new Map();
  rows.forEach((row, index) => {
    const amount = parseFlexibleNumber(row.Valor_Mercancia, null);
    if (amount === null) return;
    const date = parseSpreadsheetDate(row.Fecha_Emision);
    if (!date) return;
    const prefix = cleanText(row.Prefijo) || "FV";
    const number = cleanText(row.Numero) || cleanText(row.Numero_Documento) || String(index + 1);
    const client = cleanText(row.Empresa || row.Identificacion || "Sin cliente");
    const key = [client, prefix, number, date].join("|");
    const current =
      buckets.get(key) || {
        fechaIso: date,
        sign: amount < 0 || prefix.toUpperCase().startsWith("NC") ? -1 : 1,
        total: 0,
        totalOriginal: 0,
        cost: 0,
        categoryTotals: new Map(),
        document: [prefix, number].filter(Boolean).join("-"),
        client,
      };
    current.total += amount;
    current.totalOriginal += Math.abs(amount);
    if (amount >= 0) current.cost += parseFlexibleNumber(row.Costo_Producto, 0) || 0;
    const category = normalizeCategory(row.Categoria);
    current.categoryTotals.set(category, (current.categoryTotals.get(category) || 0) + Math.abs(amount));
    if (amount < 0) current.sign = -1;
    buckets.set(key, current);
  });

  return [...buckets.values()].map((bucket) => ({
    ...bucket,
    category: [...bucket.categoryTotals.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "Ventas",
    categoryTotals: undefined,
  }));
}

function normalizeRows(rows) {
  if (!rows.length) return [];
  if ("fechaIso" in rows[0] || "costoProducto" in rows[0] || "signoDocumento" in rows[0]) {
    return rows.map(normalizeCachedDocument).filter((row) => row.fechaIso);
  }
  return normalizeRawPbiRows(rows);
}

function summarizeDocuments(documents, range) {
  const selected = documents.filter((row) => row.fechaIso >= range.startDate && row.fechaIso <= range.endDate);
  const positives = selected.filter((row) => row.sign >= 0);
  const credits = selected.filter((row) => row.sign < 0);
  const categoryMap = new Map();

  selected.forEach((row) => {
    const category = row.category || "Ventas";
    if (!categoryMap.has(category)) {
      categoryMap.set(category, {
        category,
        grossSales: 0,
        creditNotes: 0,
        netSales: 0,
        sales: 0,
        cost: 0,
        pbiCost: 0,
        grossProfit: 0,
        marginPct: 0,
        documents: 0,
        creditDocuments: 0,
      });
    }
    const bucket = categoryMap.get(category);
    bucket.netSales += row.total;
    bucket.sales = bucket.netSales;
    if (row.sign >= 0) {
      bucket.grossSales += row.totalOriginal;
      bucket.pbiCost += row.cost;
      bucket.documents += 1;
    } else {
      bucket.creditNotes += row.totalOriginal;
      bucket.creditDocuments += 1;
    }
  });

  const categories = [...categoryMap.values()]
    .map((row) => ({
      ...row,
      ...getFinanceMarginConfig(row.category),
    }))
    .map((row) => ({
      ...row,
      sales: row.netSales,
      grossProfit: row.netSales * row.marginPct,
      cost: row.netSales - row.netSales * row.marginPct,
    }))
    .sort((a, b) => b.netSales - a.netSales);

  const grossSales = positives.reduce((sum, row) => sum + row.totalOriginal, 0);
  const creditNotes = credits.reduce((sum, row) => sum + row.totalOriginal, 0);
  const netSales = selected.reduce((sum, row) => sum + row.total, 0);
  const pbiCost = positives.reduce((sum, row) => sum + row.cost, 0);
  const grossProfit = categories.reduce((sum, row) => sum + row.grossProfit, 0);
  const cost = netSales - grossProfit;

  return {
    rows: selected.length,
    positiveDocuments: positives.length,
    creditDocuments: credits.length,
    totals: {
      grossSales,
      creditNotes,
      netSales,
      cost,
      pbiCost,
      grossProfit,
      marginPct: netSales ? grossProfit / netSales : 0,
    },
    assumptions: {
      marginSource: "Felipe.xlsx",
      fallbackMarginCategory: FINANCE_FALLBACK_MARGIN_CATEGORY,
      categoryMargins: FINANCE_CATEGORY_MARGINS,
    },
    categories,
  };
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Cache-Control", "s-maxage=900, stale-while-revalidate=1800");
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "GET") return res.status(405).json({ ok: false, error: "Metodo no permitido" });

  try {
    const range = normalizeDateRange(req);
    const source = await loadRows();
    const documents = normalizeRows(source.rows);
    const summary = summarizeDocuments(documents, range);
    return res.json({
      ok: true,
      ...range,
      source: source.source,
      sourceName: source.meta?.sourceName || (source.source === "cache" ? "Cache PBI ventas" : "API ventas PBI"),
      sourceGeneratedAt: source.generatedAt,
      sourceRange: source.meta?.range || null,
      totalAvailableDocuments: documents.length,
      ...summary,
    });
  } catch (error) {
    console.error("[finance-summary]", error);
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : "No fue posible calcular el resumen financiero.",
    });
  }
}
