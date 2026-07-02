const { corsResponse, optionsResponse } = require('./cors');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return optionsResponse();
  if (event.httpMethod !== 'POST') return corsResponse(405, 'Méthode non autorisée.');

  const SUPABASE_URL = process.env.REPORTING_SUPABASE_URL;
  const SUPABASE_KEY = process.env.REPORTING_SUPABASE_SERVICE_KEY;

  let payload;
  try { payload = JSON.parse(event.body); } catch (e) { return corsResponse(400, 'JSON invalide.'); }

  const { commercial } = payload;
  if (!commercial) return corsResponse(400, 'Paramètre "commercial" requis.');

  await fetch(`${SUPABASE_URL}/rest/v1/outlook_tokens?commercial=eq.${encodeURIComponent(commercial)}`, {
    method: 'DELETE',
    headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` }
  });

  return corsResponse(200, { ok: true });
};
