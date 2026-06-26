// netlify/functions/outlook-status.js
// Vérifie si un commercial a déjà connecté son compte Outlook.

exports.handler = async (event) => {
  const SUPABASE_URL = process.env.REPORTING_SUPABASE_URL;
  const SUPABASE_KEY = process.env.REPORTING_SUPABASE_SERVICE_KEY;

  const commercial = event.queryStringParameters && event.queryStringParameters.commercial;
  if (!commercial) {
    return { statusCode: 400, body: 'Paramètre "commercial" manquant.' };
  }

  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/outlook_tokens?commercial=eq.${encodeURIComponent(commercial)}&select=commercial,updated_at`,
    { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } }
  );
  const rows = await res.json();

  return {
    statusCode: 200,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ connected: !!(rows && rows.length), updated_at: rows && rows[0] ? rows[0].updated_at : null })
  };
};
