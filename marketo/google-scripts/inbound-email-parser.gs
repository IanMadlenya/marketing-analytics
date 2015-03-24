/***************
 * This code executes processSalesEmails() and processInboundEmails().
 * This looks at the runner's Gmail account for the labels
 * 'sales-inbound' and 'inbound-email', and parses the latest emails
 * to insert them into Marketo.
 ***************/

var FREE_DOMAINS = ['gmail','hotmail','yahoo','qq.com','msn','outlook','me.com'];
var IGNORED_EMAILS = ['info@','noreply@','contact@','@elastic.co'];
var IGNORED_BODIES = ['opt out','unsubscribe','opt-out','bank detail',' seo ',' seo.',' seo-','internship','social-media'];

function processSalesEmails() {
  var ss = getOrCreateSpreadsheet('Sales Emails');
  var sheet = ss.getSheetByName('Sent To SFDC');
  var alreadyProcessedSheet = ss.getSheetByName('Already Processed');
  var processedIds = getAlreadyProcessedEmailIds(alreadyProcessedSheet);
  var label = GmailApp.getUserLabelByName('sales-inbound');
  var threads = label.getThreads();
  var done = false;
  var toUpdate = [];
  for(var i in threads) {
    var messages = threads[i].getMessages();
    for(var c in messages) {
      var id = messages[c].getId();
      if(processedIds.indexOf(id) >= 0) {
        done = true;
        break;
      }
      alreadyProcessedSheet.appendRow([id]);
      var fromEmail = messages[c].getFrom();
      if(shouldIgnoreEmail(fromEmail.toLowerCase(),IGNORED_EMAILS)) { break; }
      var toEmail = messages[c].getTo();
      if(!toEmail ||
         (toEmail.indexOf('sales@elasticsearch.com') == -1 &&
          toEmail.indexOf('sales@elasticsearch.org') == -1 &&
          toEmail.indexOf('sales@elastic.co') == -1)
      ) { break; }
      if(toEmail.split(',').length > 2) { break; } // if it is sent to multiple addresses, it is spam
      var message;
      try {
        message = messages[c].getPlainBody();
      } catch(e) {
        break;
      }
      if(!message) { break; }
      message = message.trim();
      if(shouldIgnoreBody(message.toLowerCase(),IGNORED_BODIES)) { break; }
      var subject = messages[c].getSubject();
      var fromName = "";
      if(fromEmail.indexOf('<') >= 0) {
        fromName = fromEmail.split('<')[0].trim();
        fromName = fromName.replace(/["']/g,'');
        fromEmail = fromEmail.split('<')[1].split('>')[0].trim();
      }
      var firstName = fromName;
      var lastName = "";
      if(fromName && fromName.indexOf(' ') >= 1) {
        firstName = fromName.split(' ')[0].trim();
        lastName = fromName.replace(firstName,'').trim();
      }
      sheet.appendRow([id,fromEmail,fromName,toEmail,subject,message])
      var addWebsite = true;
      for(var x in FREE_DOMAINS) {
        if(fromEmail.toLowerCase().indexOf(FREE_DOMAINS[x]) >= 0) {
          addWebsite = false;
          break;
        }
      }
      toUpdate.push({'email':fromEmail,
                     'firstName': firstName,
                     'Form_Message':"Subject: "+subject+"\n\n"+message,
                     'lastName': lastName,
                     'Form_Source':"Inbound Email",
                     'website':(addWebsite) ? fromEmail.split('@')[1] : '' })
      break;
    }
    if(done) { break; }
  }
  updateLeads(toUpdate);
  Logger.log(ss.getUrl());
}

function processInboundEmails() {
  var ss = getOrCreateSpreadsheet('Inbound Emails');
  var sheet = ss.getSheetByName('Sent To SFDC');
  var alreadyProcessedSheet = ss.getSheetByName('Already Processed');
  var processedIds = getAlreadyProcessedEmailIds(alreadyProcessedSheet);
  var label = GmailApp.getUserLabelByName('inbound-email');
  var threads = label.getThreads();
  var done = false;
  var toUpdate = [];
  for(var i in threads) {
    var messages = threads[i].getMessages();
    for(var c in messages) {
      var id = messages[c].getId();
      if(processedIds.indexOf(id) >= 0) {
        done = true;
        break;
      }
      alreadyProcessedSheet.appendRow([id]);
      var fromEmail = messages[c].getFrom();
      if(shouldIgnoreEmail(fromEmail.toLowerCase(),IGNORED_EMAILS)) { break; }
      var toEmail = messages[c].getTo();
      if(!toEmail || (toEmail.indexOf('info@elasticsearch.com') == -1 && toEmail.indexOf('info@elasticsearch.org') == -1 && toEmail.indexOf('info@elastic.co') == -1)) { break; }
      if(toEmail.split(',').length > 2) { break; }
      var message;
      try {
        message = messages[c].getPlainBody();
      } catch(e) {
        break;
      }
      if(!message) { break; }
      message = message.trim();
      if(shouldIgnoreBody(message.toLowerCase(),IGNORED_BODIES)) { break; }
      var subject = messages[c].getSubject();
      var fromName = "";
      if(fromEmail.indexOf('<') >= 0) {
        fromName = fromEmail.split('<')[0].trim();
        fromName = fromName.replace(/["']/g,'');
        fromEmail = fromEmail.split('<')[1].split('>')[0].trim();
      }
      var firstName = fromName;
      var lastName = "";
      if(fromName && fromName.indexOf(' ') >= 1) {
        firstName = fromName.split(' ')[0].trim();
        lastName = fromName.replace(firstName,'').trim();
      }
      sheet.appendRow([id,fromEmail,fromName,toEmail,subject,message])
      var addWebsite = true;
      for(var x in FREE_DOMAINS) {
        if(fromEmail.toLowerCase().indexOf(FREE_DOMAINS[x]) >= 0) {
          addWebsite = false;
          break;
        }
      }
      toUpdate.push({'email':fromEmail,
                     'firstName': firstName,
                     'Form_Message':"Subject: "+subject+"\n\n"+message,
                     'lastName': lastName,
                     'Form_Source':"Inbound Email",
                     'website':(addWebsite) ? fromEmail.split('@')[1] : '' })
      break;
    }
    if(done) { break; }
  }
  updateLeads(toUpdate);
  Logger.log(ss.getUrl());
}

function shouldIgnoreEmail(fromEmail,ignoredEmails) {
  for(var x in ignoredEmails) {
    if(fromEmail.indexOf(ignoredEmails[x])>=0) {
      return true;
    }
  }
  return false;
}

function shouldIgnoreBody(body,ignoredBodies) {
  for(var x in ignoredBodies) {
    if(body.indexOf(ignoredBodies[x])>=0) {
      return true;
    }
  }
  return false;
}

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
      } else {
        retries--;
        Utilities.sleep(1000);
      }
    } catch (e) {
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

function getOrCreateSpreadsheet(spreadsheetName) {
  var ss;
  var files = DriveApp.getFilesByName(spreadsheetName);
  if(files.hasNext()) {
    var file = files.next();
    var retries = 3;
    while(retries > 0) {
      try {
        ss = SpreadsheetApp.openById(file.getId());
        break;
      } catch(e) {
        Utilities.sleep(1000);
        retries -= 1;
        if(retries == 0) { throw e; }
      }
    }
  } else {
    ss = SpreadsheetApp.create(spreadsheetName);
    ss.getActiveSheet().appendRow(['Id','From Email','From Name','To Email','Subject','Message'])
    ss.getActiveSheet().setName('Sent To SFDC');
    ss.insertSheet('Already Processed');
  }
  return ss;
}

function getAlreadyProcessedEmailIds(alreadyProcessedSheet) {
  var processedIds = [];
  var rawData = alreadyProcessedSheet.getRange('A:A').getValues();
  var processedIds = [];
  for(var i in rawData) {
    for(var c in rawData[i]) {
      if(rawData[i][c]) {
        processedIds.push(rawData[i][c]);
      }
    }
  }
  return processedIds;
}
