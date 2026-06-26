// netlify/functions/outlook-disconnect.js
// Supprime les tokens stockés d'un commercial (déconnexion d'Outlook).

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Méthode non autorisée.' };
  }

  const SUPABASE_URL = process.env.REPORTING_SUPABASE_URL;
  const SUPABASE_KEY = process.env.REPORTING_SUPABASE_SERVICE_KEY;

  let payload;
  try {
    payload = JSON.parse(event.body);
  } catch (e) {
    return { statusCode: 400, body: 'JSON invalide.' };
  }

  const { commercial } = payload;
  if (!commercial) {
    return { statusCode: 400, body: 'Paramètre "commercial" requis.' };
  }

  await fetch(`${SUPABASE_URL}/rest/v1/outlook_tokens?commercial=eq.${encodeURIComponent(commercial)}`, {
    method: 'DELETE',
    headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` }
  });

  return { statusCode: 200, body: JSON.stringify({ ok: true }) };
};
