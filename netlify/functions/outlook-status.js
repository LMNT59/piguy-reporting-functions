const { corsResponse, optionsResponse } = require('./cors');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return optionsResponse();

  const SUPABASE_URL = process.env.REPORTING_SUPABASE_URL;
  const SUPABASE_KEY = process.env.REPORTING_SUPABASE_SERVICE_KEY;

  const commercial = event.queryStringParameters && event.queryStringParameters.commercial;
  if (!commercial) return corsResponse(400, 'Paramètre "commercial" manquant.');

  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/outlook_tokens?commercial=eq.${encodeURIComponent(commercial)}&select=commercial,updated_at`,
    { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } }
  );
  const rows = await res.json();
  return corsResponse(200, { connected: !!(rows && rows.length), updated_at: rows && rows[0] ? rows[0].updated_at : null });
};
