const { corsResponse, optionsResponse } = require('./cors');
const { getValidAccessToken } = require('./_refresh-token');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return optionsResponse();
  if (event.httpMethod !== 'POST') return corsResponse(405, 'Méthode non autorisée.');

  let payload;
  try { payload = JSON.parse(event.body); } catch (e) { return corsResponse(400, 'JSON invalide.'); }

  const { commercial, outlook_event_id } = payload;
  if (!commercial || !outlook_event_id) {
    return corsResponse(400, 'Paramètres "commercial" et "outlook_event_id" requis.');
  }

  const accessToken = await getValidAccessToken(commercial);
  if (!accessToken) {
    return corsResponse(401, 'Ce commercial n\'a pas connecté son compte Outlook.');
  }

  const graphRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/events/${outlook_event_id}`,
    {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${accessToken}` }
    }
  );

  // 204 = suppression réussie, 404 = déjà supprimé (les deux sont OK)
  if (graphRes.status === 204 || graphRes.status === 404) {
    return corsResponse(200, { ok: true });
  }

  const errText = await graphRes.text();
  return corsResponse(graphRes.status, `Erreur Microsoft Graph : ${errText}`);
};
