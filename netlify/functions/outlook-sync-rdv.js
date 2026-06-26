// netlify/functions/outlook-sync-rdv.js
// Crée ou met à jour un événement dans le calendrier Outlook du commercial,
// à partir d'un RDV planifié dans le CRM.

const { getValidAccessToken } = require('./_refresh-token');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Méthode non autorisée.' };
  }

  let payload;
  try {
    payload = JSON.parse(event.body);
  } catch (e) {
    return { statusCode: 400, body: 'JSON invalide.' };
  }

  const { commercial, rdv } = payload;
  if (!commercial || !rdv) {
    return { statusCode: 400, body: 'Paramètres "commercial" et "rdv" requis.' };
  }

  const accessToken = await getValidAccessToken(commercial);
  if (!accessToken) {
    return { statusCode: 401, body: 'Ce commercial n\'a pas (ou plus) connecté son compte Outlook.' };
  }

  // Construit l'événement au format attendu par Microsoft Graph
  const eventBody = {
    subject: `${rdv.nom_cp || 'RDV client'}${rdv.nature_cp ? ' — ' + rdv.nature_cp : ''}`,
    body: {
      contentType: 'text',
      content: rdv.objet || ''
    },
    start: {
      dateTime: `${rdv.date_rdv}T${rdv.heure_rdv || '09:00'}:00`,
      timeZone: 'Europe/Paris'
    },
    end: {
      dateTime: `${rdv.date_rdv}T${rdv.heure_fin_rdv || '10:00'}:00`,
      timeZone: 'Europe/Paris'
    },
    location: { displayName: rdv.nom_cp || '' }
  };

  const graphUrl = rdv.outlook_event_id
    ? `https://graph.microsoft.com/v1.0/me/events/${rdv.outlook_event_id}`
    : `https://graph.microsoft.com/v1.0/me/events`;

  const method = rdv.outlook_event_id ? 'PATCH' : 'POST';

  const graphRes = await fetch(graphUrl, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(eventBody)
  });

  if (!graphRes.ok) {
    const errText = await graphRes.text();
    return { statusCode: graphRes.status, body: `Erreur Microsoft Graph : ${errText}` };
  }

  const createdEvent = await graphRes.json();

  return {
    statusCode: 200,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ outlook_event_id: createdEvent.id })
  };
};
