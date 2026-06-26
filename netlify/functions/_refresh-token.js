// netlify/functions/_refresh-token.js
// Fonction utilitaire (pas exposée directement) pour renouveler un token expiré.
// Utilisée par outlook-sync-rdv.js avant chaque appel à Microsoft Graph.

async function getValidAccessToken(commercial) {
  const SUPABASE_URL = process.env.REPORTING_SUPABASE_URL;
  const SUPABASE_KEY = process.env.REPORTING_SUPABASE_SERVICE_KEY;
  const CLIENT_ID = process.env.MS_CLIENT_ID;
  const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
  const TENANT_ID = process.env.MS_TENANT_ID;

  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/outlook_tokens?commercial=eq.${encodeURIComponent(commercial)}&select=*`,
    { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } }
  );
  const rows = await res.json();
  if (!rows || !rows.length) return null;
  const row = rows[0];

  const stillValid = new Date(row.expires_at).getTime() > Date.now() + 60000;
  if (stillValid) return row.access_token;

  // Token expiré : on le renouvelle avec le refresh_token
  const refreshRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      grant_type: 'refresh_token',
      refresh_token: row.refresh_token,
      scope: 'offline_access User.Read Calendars.ReadWrite'
    })
  });

  if (!refreshRes.ok) return null;
  const tokens = await refreshRes.json();
  const expiresAt = new Date(Date.now() + tokens.expires_in * 1000).toISOString();

  await fetch(`${SUPABASE_URL}/rest/v1/outlook_tokens?commercial=eq.${encodeURIComponent(commercial)}`, {
    method: 'PATCH',
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${SUPABASE_KEY}`,
      'Content-Type': 'application/json',
      Prefer: 'return=minimal'
    },
    body: JSON.stringify({
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token || row.refresh_token,
      expires_at: expiresAt,
      updated_at: new Date().toISOString()
    })
  });

  return tokens.access_token;
}

module.exports = { getValidAccessToken };
