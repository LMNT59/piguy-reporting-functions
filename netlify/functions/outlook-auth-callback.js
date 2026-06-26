// netlify/functions/outlook-auth-callback.js
// Microsoft redirige ici après que le commercial a autorisé l'accès.
// On échange le "code" reçu contre un vrai token d'accès, puis on le stocke dans Supabase.

exports.handler = async (event) => {
  const CLIENT_ID = process.env.MS_CLIENT_ID;
  const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
  const TENANT_ID = process.env.MS_TENANT_ID;
  const REDIRECT_URI = process.env.MS_REDIRECT_URI;
  const SUPABASE_URL = process.env.REPORTING_SUPABASE_URL;
  const SUPABASE_KEY = process.env.REPORTING_SUPABASE_SERVICE_KEY; // clé service_role, jamais la clé anon ici

  const { code, state, error, error_description } = event.queryStringParameters || {};

  if (error) {
    return { statusCode: 400, body: `Erreur Microsoft : ${error_description || error}` };
  }
  if (!code || !state) {
    return { statusCode: 400, body: 'Paramètres manquants (code ou state).' };
  }

  const commercial = decodeURIComponent(state);

  // Échange du code contre les tokens
  const tokenRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      grant_type: 'authorization_code',
      code,
      redirect_uri: REDIRECT_URI,
      scope: 'offline_access User.Read Calendars.ReadWrite'
    })
  });

  if (!tokenRes.ok) {
    const errText = await tokenRes.text();
    return { statusCode: 500, body: `Erreur lors de l'échange du token : ${errText}` };
  }

  const tokens = await tokenRes.json();
  const expiresAt = new Date(Date.now() + tokens.expires_in * 1000).toISOString();

  // Stocke (ou met à jour) les tokens du commercial dans Supabase
  const upsertRes = await fetch(`${SUPABASE_URL}/rest/v1/outlook_tokens?on_conflict=commercial`, {
    method: 'POST',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${SUPABASE_KEY}`,
      'Content-Type': 'application/json',
      'Prefer': 'resolution=merge-duplicates'
    },
    body: JSON.stringify({
      commercial,
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      expires_at: expiresAt,
      updated_at: new Date().toISOString()
    })
  });

  if (!upsertRes.ok) {
    const errText = await upsertRes.text();
    return { statusCode: 500, body: `Erreur lors de la sauvegarde du token : ${errText}` };
  }

  // Redirige le commercial vers le CRM avec un message de succès
  const CRM_URL = process.env.CRM_BASE_URL || '/';
  return {
    statusCode: 302,
    headers: { Location: `${CRM_URL}?outlook=connecte` },
    body: ''
  };
};
