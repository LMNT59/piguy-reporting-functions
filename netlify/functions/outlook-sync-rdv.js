const { corsResponse, optionsResponse } = require('./cors');
const { getValidAccessToken } = require('./_refresh-token');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return optionsResponse();
  if (event.httpMethod !== 'POST') return corsResponse(405, 'Méthode non autorisée.');

  let payload;
  try { payload = JSON.parse(event.body); } catch (e) { return corsResponse(400, 'JSON invalide.'); }

  const { commercial, rdv } = payload;
  if (!commercial || !rdv) return corsResponse(400, 'Paramètres "commercial" et "rdv" requis.');

  const accessToken = await getValidAccessToken(commercial);
  if (!accessToken) return corsResponse(401, 'Ce commercial n\'a pas connecté son compte Outlook.');

  const eventBody = {
    subject: `${rdv.nom_cp || 'RDV client'}${rdv.nature_cp ? ' — ' + rdv.nature_cp : ''}`,
    body: { contentType: 'text', content: rdv.objet || '' },
    start: { dateTime: `${rdv.date_rdv}T${rdv.heure_rdv || '09:00'}:00`, timeZone: 'Europe/Paris' },
    end: { dateTime: `${rdv.date_rdv}T${rdv.heure_fin_rdv || '10:00'}:00`, timeZone: 'Europe/Paris' },
    location: { displayName: rdv.nom_cp || '' }
  };

  const graphUrl = rdv.outlook_event_id
    ? `https://graph.microsoft.com/v1.0/me/events/${rdv.outlook_event_id}`
    : `https://graph.microsoft.com/v1.0/me/events`;

  const method = rdv.outlook_event_id ? 'PATCH' : 'POST';

  const graphRes = await fetch(graphUrl, {
    method,
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(eventBody)
  });

  if (!graphRes.ok) {
    const errText = await graphRes.text();
    return corsResponse(graphRes.status, `Erreur Microsoft Graph : ${errText}`);
  }

  const createdEvent = await graphRes.json();
  return corsResponse(200, { outlook_event_id: createdEvent.id });
};
