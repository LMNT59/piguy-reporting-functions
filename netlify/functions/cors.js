const CORS_HEADERS = {
  'Access-Control-Allow-Origin': 'https://reportingpiguy.netlify.app',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

function corsResponse(statusCode, body) {
  return {
    statusCode,
    headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' },
    body: typeof body === 'string' ? body : JSON.stringify(body)
  };
}

function optionsResponse() {
  return { statusCode: 204, headers: CORS_HEADERS, body: '' };
}

module.exports = { CORS_HEADERS, corsResponse, optionsResponse };
