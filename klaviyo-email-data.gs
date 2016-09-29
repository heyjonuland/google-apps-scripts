// Set up a trigger to pull this data once/day or once/week.

function getEmailMetrics() {
  //Dates
  var d = new Date();
  var today = Utilities.formatDate(d, "ET", "yyyy-MM-dd"); // Klaviyo-accepted date format
  
  // Function that returns a date 'n' days before the input date
  function dateBefore(date,n){
    var result = Utilities.formatDate(new Date(date.getTime()-n*(24*3600*1000)), "ET", "yyyy-MM-dd");
    return result
  }  
  
  //API Request
  var options = 
      {
        "method" : "get"
      };
  var url = 'https://a.klaviyo.com/api/v1/metric/';
  var token = '{{ KLAVIYO TOKEN }}';
  var startDate =  dateBefore(d,7); // Sets the start date to 7 days before today.
  var endDate = today;
  var unit = 'week'; // Klaviyo params
  var measurement = 'value'; // Klaviyo params
  

  // Received  
  var receivedId = '{{ RECEIVED ID }}';
  received = UrlFetchApp.fetch(url + receivedId + '/export' + '?' + 'api_key=' + token + "&start_date=" + startDate + "&end_date=" + endDate + "&unit=" + unit, options);
  
  var blob = JSON.parse(received);
  var received = blob.results[0].data[0].values[0];
  Logger.log("Received: " + received);
  
  // Opened
  var openedId = '{{ OPENED ID }}';
  opened = UrlFetchApp.fetch(url + openedId + '/export' + '?' + 'api_key=' + token + "&start_date=" + startDate + "&end_date=" + endDate + "&unit=" + unit, options);
  
  var blob = JSON.parse(opened);
  var opens = blob.results[0].data[0].values[0];
  if (received == 0) { var openRate = 0 } else { var openRate = opens / received};
  Logger.log("Opened: " + opens);
  
  // Clicked
  var clickedId = '{{ CLICKED ID }}';
  clicked = UrlFetchApp.fetch(url + clickedId + '/export' + '?' + 'api_key=' + token + "&start_date=" + startDate + "&end_date=" + endDate + "&unit=" + unit, options);
  
  var blob = JSON.parse(clicked);
  var clicks = blob.results[0].data[0].values[0];
  if (received == 0) { var clickRate = 0 } else { var clickRate = clicks / received};
  Logger.log("Clicked: " + clicks);
  
  // Unsubscribed
  var unsubId = '{{ UNSUBSCRIBED ID }}';
  unsubscribed = UrlFetchApp.fetch(url + unsubId + '/export' + '?' + 'api_key=' + token + "&start_date=" + startDate + "&end_date=" + endDate + "&unit=" + unit, options);
  
  var blob = JSON.parse(unsubscribed);
  var unsubs = blob.results[0].data[0].values[0];
  if (received == 0) { var unsubRate = 0 } else { var unsubRate = unsubs / received};
  Logger.log("Unsubscribed: " + unsubs);

  // Revenue
  var revenueId = '{{ STRIPE SUCCESSFUL PAYMENT ID }}';
  revenue = UrlFetchApp.fetch(url + revenueId + '/export' + '?' + 'api_key=' + token + "&start_date=" + startDate + "&end_date=" + endDate + "&unit=" + unit + "&measurement=" + measurement, options);
  
  var blob = JSON.parse(revenue);
  var revenue = blob.results[0].data[0].values[0];
  Logger.log("Revenue: " + revenue);
  
  // Refunded Payment
  var refundId = '{{ STRIPE REFUND PAYMENT ID }}';
  refund = UrlFetchApp.fetch(url + refundId + '/export' + '?' + 'api_key=' + token + "&start_date=" + startDate + "&end_date=" + endDate + "&unit=" + unit + "&measurement=" + measurement, options);
  
  var blob = JSON.parse(refund);
  var refund = blob.results[0].data[0].values[0];
  Logger.log("Refund: " + refund);

  // Spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([today,received,opens,clicks,unsubs,openRate,clickRate,unsubRate,revenue,refund, (revenue - refund)]);
  
}
