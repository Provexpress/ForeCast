export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  res.setHeader('Cache-Control', 's-maxage=3600'); // cache 1 hora

  try {
    const r = await fetch('https://www.dolar-colombia.com/', {
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    const html = await r.text();
    const m = html.match(/1\s+USD\s*=\s*([\d,\.]+)\s*COP/i);
    if (m) {
      const trm = parseFloat(m[1].replace(/,/g, ''));
      return res.json({ trm, source: 'dolar-colombia.com', ok: true });
    }
    throw new Error('No se encontró el valor');
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
}
