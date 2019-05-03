// This function will fetch the following metrics in this order (so be sure to add titles to these columns!): 
// Sent |	Delivered |	Total Opens |	Unique Opens |	Total Clicks |	Unique Clicks |	Unsubscribes |	Bounces |	Conversions (or Purchases)

var token = "YOUR ITERABLE API KEY";
var numHeaderRows = 1

// IF COLUMNS ADDED/REMOVED, UPDATE THESE VARS
// "cid" = "campaign ID" --> This can be found in the URL of each campaign, or in the campaign table view in Iterable.
// "One Base" and "Zero Base" refer to counting methods. "One Base" means the first column number is 1, i.e. column A = column 1.

var cidColOneBase = 3
var lockColOneBase = 4
var startDateOneBase = 11
var endDateOneBase = 12
var metricsStartColOneBase = 17

//
////
////// DO NOT CHANGE BELOW THIS LINE (unless you absolutely know what you're doing)

var cidColZeroBase = cidColOneBase - 1
var lockColZeroBase = lockColOneBase - 1
var startDateZeroBase = startDateOneBase - 1
var endDateZeroBase = endDateOneBase - 1
var metricsStartColZeroBase = metricsStartColOneBase - 1

// This function is what puts everything together and requests metrics from Iterable
function getIterableMetrics() {
  
  // get the data range
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  var startRow = numHeaderRows + 1
  
  // get the number of rows in the data range
  var campRows = campaignRows(numHeaderRows,cidColZeroBase,metricsStartColZeroBase)
  
  var lastRow = campRows.full + campRows.empty
  var lastRowIndex = lastRow + 1
  var r = startRow
  
  while (r <= lastRowIndex) {
    
    //   get the values needed to run metrics function
    var campaignId = sheet.getRange(r, cidColOneBase).getValue()
    var startDate = sheet.getRange(r, startDateOneBase).getValue()
    var endDate = sheet.getRange(r, endDateOneBase).getValue()
    var metricsCol = sheet.getRange(r, metricsStartColOneBase).getValue()
    var lockCol = sheet.getRange(r,lockColOneBase).getValue()
    
    // Format dates
    var s = Utilities.formatDate(new Date(startDate),"America/New York","yyyy-MM-dd hh:mm:ss")
    var e = Utilities.formatDate(new Date(endDate.getTime() + 1*(24*3600*1000)),"America/New York","yyyy-MM-dd hh:mm:ss");  
    
    // Set up output blob
    var metricsBlob = []
    
    //   if data columns are blank 
    if (typeof(metricsCol) != "number") {
      //     run the metrics function
      var response = getMetrics(campaignId,s,e)
      // Logger.log(response.length)
      
      // Set values in sheet
      sheet.getRange(r, metricsStartColOneBase, 1, response.length).setValues([response])
      
      
    } else {
      //     if lock metrics = false
      if (lockCol == false) {
        //       run the metrics function
        var response = getMetrics(campaignId,s,e)
        // Logger.log(response.length)
        
        // Set values in sheet
        sheet.getRange(r, metricsStartColOneBase, 1, response.length).setValues([response])
      
      } else {
        // do nothing
        Logger.log("Doing nothing")
      }
    }
    r++
  }
  
}

// This is the actual function that talks to Iterable
function getMetrics(cid,s,e) {
  
  // Build endpoint URL
  var url = "https://api.iterable.com/api/campaigns/metrics?campaignId=" + cid + "&startDateTime=" + s + "&endDateTime=" + e + "&api_key=" + token;
  
  var options = {
    "method" : "get",
    'contentType': 'application/json',
    "muteHttpExceptions" : false
  };
  
  // Logger.log("about to try api get");
  
  // Fetch Response
  try {
    var response = UrlFetchApp.fetch(url, options);
    // Logger.log(response)
  }
  catch (err) {
    Logger.log(response)
    
  }
  
  
  // Declare metrics var
  var metrics = {};
  
  // Convert response to string
  var str = response.toString();
  
  // Prep CSV for JSON Conversion
  var rows = str.split("\n");
  
  // Format Headers
  var headers = rows[0].toLowerCase().replace(/\//g,"by").replace(/\s[mM]/g,"_mil").replace(/\s/g, "_").split(",");
  // Logger.log(headers)
  
  // Create Values
  var values = rows[1].split(",");
  
  // Create JSON Object
  for (i = 0; i < headers.length; i++) {
    metrics[headers[i]] = values[i]
  } 
  
  // AVAILABLE METRICS
  var id = metrics.id;
  var average_order_value = JSON.parse(metrics.average_order_value);
  var purchases_by_mil = JSON.parse(metrics.purchases_by_mil);
  var revenue = JSON.parse(metrics.revenue);
  var revenue_by_mil = JSON.parse(metrics.revenue_by_mil);
  var total_complaints = JSON.parse(metrics.total_complaints);
  var total_email_opens = JSON.parse(metrics.total_email_opens);
  var sent = JSON.parse(metrics.total_email_sends);
  var total_emails_bounced = JSON.parse(metrics.total_emails_bounced);
  var total_emails_clicked = JSON.parse(metrics.total_emails_clicked);
  var delivered = JSON.parse(metrics.total_emails_delivered);
  var total_hosted_unsubscribe_clicks = JSON.parse(metrics.total_hosted_unsubscribe_clicks);
  var total_unsubscribes = JSON.parse(metrics.total_unsubscribes);
  var unique_email_clicks = JSON.parse(metrics.unique_email_clicks);
  var unique_opens = JSON.parse(metrics.unique_email_opens_or_clicks);
  var unique_emails_bounced = JSON.parse(metrics.unique_emails_bounced);
  var unique_hosted_unsubscribe_clicks = JSON.parse(metrics.unique_hosted_unsubscribe_clicks);
  var unique_purchases = JSON.parse(metrics.unique_purchases);
  var unique_unsubscribes = JSON.parse(metrics.unique_unsubscribes);
  if (metrics.unique_custom_conversions) { 
    var conversions = JSON.parse(metrics.unique_custom_conversions) 
    } else { 
      var conversions = JSON.parse(metrics.total_purchases)
      };
  
  // Prepare function output
  var output = [sent,delivered,total_email_opens,unique_opens,total_emails_clicked,unique_email_clicks,unique_hosted_unsubscribe_clicks,unique_emails_bounced,conversions]
  // Adds a 300 millisecond delay to handle date changes and to try and curb rate limits
  Utilities.sleep(300); 
  
  // Logger.log(output)
  return output
  
}

function getListSize(listId) {
  
  // Build endpoint URL
  var url = "https://api.iterable.com/api/lists/" + listId + "/size?" + "api_key=" + token;
  
  var options = {
    "method" : "get",
    'contentType': 'application/json',
    "muteHttpExceptions" : true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
  } catch (err) {
    throw ("Yikes! Something went wrong :( Wait a couple of minutes, change the dates, and try again.");
  }
  
  Logger.log(response)
  
  var output = [response]
  
  Logger.log(output)
  return JSON.parse(response)
  
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('Get Metrics')
  .addItem('Get Metrics from Iterable', 'getIterableMetrics')
  .addToUi();
}
