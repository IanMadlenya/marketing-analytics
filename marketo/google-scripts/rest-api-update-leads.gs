/****************************
 * Calls the Marketo rest api to update leads.
 * Requires: rest-api-auth.gs to get accessToken
 * toUpdate format looks something like this:
 * var toUpdate = [{'email':'bill@clinton.com',
 *                  'firstName': 'Bill',
 *                  'lastName': 'Clinton'}];
 *
 ****************************/
function updateLeads(toUpdate) {
  var restEndpoint = "https://813-mam-392.mktorest.com/rest/v1/leads.json";
  var accessToken = getMarketoAccessToken();
  var options = {
    method:'POST',
    payload:JSON.stringify({
      "action":"createOrUpdate",
      "lookupField":"email",
      "input":toUpdate
    }),
    contentType : 'application/json',
    headers: {
      'Authorization': 'Bearer '+accessToken,
      'Accept': 'application/json'
    },
    muteHttpExceptions:true
  };
  var retries = 10;
  var resp;
  while(retries > 0) {
    try {
      resp = UrlFetchApp.fetch(restEndpoint,options);
      if(resp.getResponseCode() == 200) {
        Logger.log(resp.getContentText());
        return;
      } else { // if we got a 400 or something
        retries--;
        Utilities.sleep(1000);
      }
    } catch (e) { // If UrlFetch fails for some reason
      retries--;
      Utilities.sleep(1000);
    }
  }
  if(resp) {
    throw 'something went wrong sending data: '+resp.getContentText();
  } else {
    throw 'something went wrong: '+JSON.stringify(toUpdate);
  }
}
