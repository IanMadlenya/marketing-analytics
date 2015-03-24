/****************************
 * Calls the Marketo identity service and returns
 * the access_token used for additional requests.
 * Usage example: see rest-api-update-leads.gs
 ****************************/
function getMarketoAccessToken() {
  var identityServiceUrl = "https://813-MAM-392.mktorest.com/identity/oauth/token";
  var grantType = "client_credentials";
  var clientId = "d42e42c5-a619-4d6f-b3ce-4fcae79fbda3";
  var clientSecret = "xsMA3me9EWYEntxWiO3oANdV9Jxl8nYp";
  var reqUrl = identityServiceUrl+'?grant_type='+grantType+'&client_id='+clientId+'&client_secret='+clientSecret;
  var resp;
  var retries = 10;
  while(retries > 0) {
    resp = UrlFetchApp.fetch(reqUrl,{muteHttpExceptions:true});
    if(resp.getResponseCode() == 200) {
      var respJson = JSON.parse(resp.getContentText());
      return respJson.access_token
    } else {
      retries--;
      Utilities.sleep(1000);
    }
  }
  throw Utilities.formatString("Something went wrong getting credentials: Request Url: %s Message: %s", reqUrl, resp.getContentText());
}
