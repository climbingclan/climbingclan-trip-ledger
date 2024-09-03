var server = '18.168.242.164'; 
var port = 3306; 
var dbName = 'bitnami_wordpress'; 
var username = 'gsheets'; 
var password = 'eyai4yohF4uX8eeP7phoob'; 
var url = 'jdbc:mysql://'+server+':'+port+'/'+dbName; 
var cc_location = "Parthian Climbing Manchester"; 

function retrieveOrderData() { 
  var conn = Jdbc.getConnection(url, username, password); 
  var stmt = conn.createStatement(); 
  var years = [2024, 2023, 2022]; 
  
  // Get the current year index from Script Properties
  var scriptProperties = PropertiesService.getScriptProperties();
  var currentYearIndex = scriptProperties.getProperty('currentYearIndex');
  
  // If the current year index is not set, initialize it to 0
  if (currentYearIndex === null) {
    currentYearIndex = 0;
  } else {
    currentYearIndex = parseInt(currentYearIndex);
  }
  
  // Get the year based on the current year index
  var year = years[currentYearIndex % years.length];
  
  console.log(year);
  
  var query = 'SELECT order_id as "Order ID", first_name as "First Name", last_name "Last Name", ' + 
              'order_item_name "Trip", order_created "Order Created" ' + 
              'FROM wp_vw_order_details ' + 
              'WHERE YEAR(order_created) = ' + year; 
              
  var results = stmt.executeQuery(query); 
  var metaData = results.getMetaData(); 
  var numCols = metaData.getColumnCount(); 
  
  var spreadsheet = SpreadsheetApp.getActive(); 
  var sheetName = 'Orders ' + year; 
  var sheet = spreadsheet.getSheetByName(sheetName); 
  
  if (sheet === null) { 
    sheet = spreadsheet.insertSheet(sheetName); 
  } else { 
    sheet.clearContents(); 
  } 
  
  var headerRow = []; 
  for (var col = 0; col < numCols; col++) { 
    headerRow.push(metaData.getColumnName(col + 1)); 
  } 
  sheet.appendRow(headerRow); 
  
  var dataRows = []; 
  while (results.next()) { 
    var dataRow = []; 
    for (var col = 0; col < numCols; col++) { 
      dataRow.push(results.getString(col + 1)); 
    } 
    dataRows.push(dataRow); 
  } 
  sheet.getRange(2, 1, dataRows.length, numCols).setValues(dataRows); 
  sheet.autoResizeColumns(1, numCols); 
  
  results.close(); 
  stmt.close(); 
  conn.close(); 
  
  // Update the current year index in Script Properties
  currentYearIndex++;
  scriptProperties.setProperty('currentYearIndex', currentYearIndex.toString());
}