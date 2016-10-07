function createNewDoc() {
  
  var templateId = '[ DOC ID FOR RESPONSE TEMPLATE ]';
  
  // Get active sheet in spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  // Compile row data into iterable object
  var startRow = 2;  // First row of data to process
  var range = sheet.getDataRange();
  var numRows = range.getNumRows() - 1;
  var numCols = range.getNumColumns();
  var dataRange = sheet.getRange(startRow, 1, numRows, numCols );
  var data = dataRange.getValues();
  
  
  // Iterate over row data and place in variables
  
  for (i in data) {
    var row = data[i];
    var timestamp = Utilities.formatDate(row[0], 'America/New_York', 'd MMMM, yyyy');
    var username = row[1];
    var title = row[2];
    var requestCategory = row[3];
    var requestType = row[4];
    var city = row[5];
    var effectiveDate = row[6];
    var startTime = Utilities.formatDate(row[7], 'America/New_York', 'hh:mm a');
    var endTime = Utilities.formatDate(row[8], 'America/New_York', 'hh:mm a');
    var locationName = row[9];
    var locationAddress = row[10];
    var memberCost = row[11];
    var guestPolicy = row[12];
    var guestCost = row[13];
    var guestsAllowed = row[14];
    var stakeholderEmail = row[15];
    var coreOffer = row[16];
    var description = row[17];
    var specialDetails = row[18];
    var heroImage = row[19];
    var audience = row[20];
    var managementCompany = row[21];
    var buildingName = row[22];
    var externalCta = row[23];
    var publishChannels = row[24];
    
    var internalStakeholders = [stakeholderEmail + "," + username];
   
  
   // Create new Document and add editor permissions for someone.
   var folder = DriveApp.getFolderById('[ FOLDER ID ]');
   var newDoc = DriveApp.getFileById(templateId).makeCopy(Utilities.formatDate(effectiveDate,'America/New_York','yyyy.MM.dd') + "_" + city + "_" + title, folder);
   var newDocId = newDoc.getId();
   newDoc.addEditors(['ADD EDITOR EMAILS HERE']);
   newDoc.addCommenter(username);
   
   // Replace placeholder text in template copy with spreadsheet values.
   var doc = DocumentApp.openById(newDocId);
   var body = doc.getActiveSection();
   body.replaceText('{{ Timestamp }}',timestamp);
   body.replaceText('{{ Username }}',username);
   body.replaceText('{{ RFC Title }}',title);
   body.replaceText('{{ Request Category }}',requestCategory);
   body.replaceText('{{ Request Type }}',requestType);
   body.replaceText('{{ City }}',city);
   body.replaceText('{{ Effective Date }}',Utilities.formatDate(effectiveDate, 'America/New_York', 'MM/dd/yyyy'));
   body.replaceText('{{ Start Time }}',startTime);
   body.replaceText('{{ End Time }}',endTime);
   body.replaceText('{{ Location Name }}',locationName);
   body.replaceText('{{ Location Address }}',locationAddress);
   body.replaceText('{{ Member Cost }}',memberCost);
   body.replaceText('{{ Guest Policy }}',guestPolicy);
   body.replaceText('{{ Guest Cost }}',guestCost);
   body.replaceText('{{ Number of Guests Allowed }}',guestsAllowed);
   body.replaceText('{{ Stakeholder Email Addresses }}',stakeholderEmail);
   body.replaceText('{{ Core Offer }}',coreOffer);
   body.replaceText('{{ Brief Description }}',description);
   body.replaceText('{{ Special Request Details }}',specialDetails);
   body.replaceText('{{ Hero Image URL }}',heroImage);
   body.replaceText('{{ Request Audience }}',audience);
   body.replaceText('{{ Management Company }}',managementCompany);
   body.replaceText('{{ Building Name }}',buildingName);
   body.replaceText('{{ External CTA URL }}',externalCta);
   body.replaceText('{{ Publishing Channels }}',publishChannels);
   doc.saveAndClose();

    // Assign Asana Tags. Prior use case involved different tags for geo-specific tasks (i.e. "New York," "San Francisco," etc.)
    var tag = null;
    if (city == "" || "") {
      tag = "TAG_ID";
    } else if ( city == "" || "") {
      tag = "TAG_ID";
    } else if ( city == "" || "") {
      tag = "TAG_ID";
    } else {
      tag = null;
    };


    // Build Asana bundle
    var workspace = 'WORKSPACE_ID'; // Workspace ID
    var project = 'PROJECT_ID'; // Project ID
    var bundle = {
      "data": {
        "workspace": workspace,
        "projects": project,
        "name": title,
        "due_on": Utilities.formatDate(effectiveDate,'EST','yyyy-MM-dd'),
        "html_notes": "", // Input HTML to create rich text
        "tags": [ tag ],
        "custom_fields": { "FIELD_ID" : "FIELD_SELECTION_ID" } // For custom field type = enum
      }
    };

    postAsana(bundle);
  }

  sheet.deleteRows(startRow, numRows);
};

function postAsana(input) {

  // Asana API
  var url = 'https://app.asana.com/api/1.0/';
  var token = 'ASANA_TOKEN';

  var payload = input;
  var options = {
    "method" : "post",
    "headers" : {
      "Accept": "application/json",        // accept JSON format
      "Content-Type": "application/json",  //content
      "Authorization": "Bearer " + token // authorization
     },
     "payload": JSON.stringify(payload)
  };

    Logger.log(options);

  var response = UrlFetchApp.fetch(url + 'tasks', options);

  Logger.log(response);

};
