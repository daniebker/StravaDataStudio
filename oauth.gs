var CLIENT_ID = 'XXX';
var CLIENT_SECRET = 'YYY';

// configure the service
function getStravaService() {
  return OAuth2.createService('Strava')
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('activity:read_all');
}

// handle the callback
function authCallback(request) {
  var stravaService = getStravaService();
  var isAuthorized = stravaService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}
