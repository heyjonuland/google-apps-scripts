// This function currently breaks down when there is an invalid email (or one that that Klaviyo rejects). 
// I have not built in error handling for this.


function updateListMembers() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var range = sheet.getDataRange();
  var numRows = range.getNumRows() - 1;
  var numCols = range.getNumColumns();
  var dataRange = sheet.getRange(startRow, 1, numRows, numCols );
  var data = dataRange.getValues();
  
  for (i in data) {
    var row = data[i];
    var email = row[1];
    var firstName = row[2]; 
    var lastName = row[3];
    var defaultCity = row[4];
    var plusMember = row[5];
    var phoneNumber = row[6];
    var streetAddress = row[7];
    var city = row[8];
    var state = row[9];
    var zipCode = row[10];
    var role = row[11].toLowerCase();
    var managementCompany = row[12];
    var buildingName = row[13];
    var gender = row[14].toLowerCase();
    var birthday = row[15];
    
    var jsonProps = {"$first_name" : firstName, 
                     "$last_name" : lastName, 
                     "gender" : gender, 
                     "default_city" : defaultCity, 
                     "role" : role, 
                     "plus_member" : plusMember, 
                     "phone_number" : phoneNumber, 
                     "birthday" : birthday, 
                     "city" : city, 
                     "state" : state, 
                     "street_address" : streetAddress, 
                     "zip_code" : zipCode, 
                     "management_company" : managementCompany, 
                     "building_name" : buildingName
                    }
    
    var bundle = { "email" : email, 
                  "properties": JSON.stringify(jsonProps),
                 "confirm_optin":"false"
               };
    
    Logger.log(bundle);
    
    var options =
        {
          "method" : "post",
          "payload" : bundle
        };
    
    var url = 'https://a.klaviyo.com/api/v1/list/';
    var listId = '{{ POST LIST ID }}';
    var token = '{{ KLAVIYO TOKEN }}';
    
    response = UrlFetchApp.fetch(url + listId + '/members' + '?' + '&api_key=' + token, options);
    
    Logger.log(response);
    
  }
  
  sheet.deleteRows(startRow, numRows); // Deletes the form submission
  
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.addMenu("Update Members", [{
    name : "Update Members",
    functionName : "updateListMembers"
  }]);
};
