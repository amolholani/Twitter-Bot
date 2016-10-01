function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Twitter Menu')
      .addItem('Send DM', 'sendFromSpreadsheet')
      .addToUi();
}

function sendFromSpreadsheet(){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getDataRange().getValues();
  
  for(var i = 1; i < data.length; i++){
    Logger.log("User  : " + data[i][0]);
    Logger.log("DM    : " + data[i][1]);
    
    var user = data[i][0];
    var msg = data[i][1];
    var status = data[i][2];
     
    var x = i + 1;
    if(status == "Not sent"){                                                                                           // if the status is "not sent", then only send the DM.
      var response = sendMsg(user, msg);
      sheet.getRange("c" + x).setValue(response);                                                                       // update the response in the spreadsheet
      Logger.log("Status: " + response);   
    }  
  }
}

function sendMsg(user, tweet) {
  
  var twitterKeys= {
    TWITTER_CONSUMER_KEY: "ZxkdSXMskCrWkXjWHmZeJaF0B",
    TWITTER_CONSUMER_SECRET: "8qIlQaRk3S8CmNBStVr1Q1Ect6gSjn7ciHMZoamvSEG2t9XXCR",
    TWITTER_ACCESS_TOKEN: "777902369854124032-iFrQMrzak5KXgZ09Rb5bRiT1hr6aV4v",
    TWITTER_ACCESS_SECRET: "TWHejuQT0bZJN4MCD37r6Y5R0E8s8hAlHev1MMrtazIFh"    
  };
  var props = PropertiesService.getScriptProperties();
  props.setProperties(twitterKeys);
  var service = new Twitter.OAuth(props);
  
  if ( service.hasAccess() ) {  
        var twitterUser = user.trim().replace(/^\@/, "");
        var api = "https://api.twitter.com/1.1/";
        api += "direct_messages/new.json?screen_name=" + twitterUser + "&text=" + encodeString_(tweet);                // Send a public direct message (DM)
        var response = service.fetch(api, {
            method: "POST",
            muteHttpExceptions: true
        });
    
    var x = JSON.parse(response);                                                                                      // parsing the response recieved by the googlescript
    
    if (x.id) {  // if sent successfully
      
      Logger.log("Tweet ID " + response.id_str);
      return "sent";
      
    } else {                                                                                                           // if error occured while sending the DM
      Logger.log(x.errors[0].message);
      return x.errors[0].message;
    }
  }  
}

function encodeString_(q) {
    var str = q;
    str = str.replace(/!/g, 'Ị');
    str = str.replace(/\*/g, '×');
    str = str.replace(/\(/g, '[');
    str = str.replace(/\)/g, ']');
    str = str.replace(/'/g, '’');
    return encodeURIComponent(str);
}
