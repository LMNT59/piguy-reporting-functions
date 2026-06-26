// netlify/functions/outlook-auth-start.js
// Démarre la connexion OAuth avec Microsoft. Redirige le commercial vers la page
// de connexion Microsoft pour qu'il autorise l'accès à son calendrier Outlook.

exports.handler = async (event) => {
  const CLIENT_ID = process.env.MS_CLIENT_ID;
  const TENANT_ID = process.env.MS_TENANT_ID;
  const REDIRECT_URI = process.env.MS_REDIRECT_URI; // ex: https://votre-site.netlify.app/.netlify/functions/outlook-auth-callback

  const commercial = event.queryStringParameters && event.queryStringParameters.commercial;
  if (!commercial) {
    return { statusCode: 400, body: 'Paramètre "commercial" manquant.' };
  }

  const scopes = [
    'offline_access',
    'User.Read',
    'Calendars.ReadWrite'
  ].join(' ');

  // "state" permet de savoir, au retour de Microsoft, quel commercial a autorisé l'accès
  const state = encodeURIComponent(commercial);

  const authUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize` +
    `?client_id=${CLIENT_ID}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&response_mode=query` +
    `&scope=${encodeURIComponent(scopes)}` +
    `&state=${state}`;

  return {
    statusCode: 302,
    headers: { Location: authUrl },
    body: ''
  };
};
