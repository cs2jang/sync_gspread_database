function getConnection(){
  // <input your infomation>
  var address = '<db host address>';
  var user = '<db user>';
  var userPwd = '<db password>';
  var db = '<db name>';
  var spreadsheetURL = "<google sheet addreess>"; 
  // <google sheet addreess> like https://docs.google.com/spreadsheets/d/10YxcTogE####g/edit#gid=0
      
  var dbUrl = 'jdbc:mysql://' + address + '/' + db;
  var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  return conn;
}

function selectData(e) {
  try {
    Logger.log('Try Read Data');
    // select current activated sheet
    var sheet = SpreadsheetApp.getActiveSheet();
    var table_name = sheet.getRange('B2').getValue();
    // make connection
    var conn = getConnection();
    var stmt = conn.createStatement();
    var query_string = 'select * from ' + table_name
    // for check if there is right data
    Logger.log('Selected Table: ' + sheet.getRange('B2').getValue());
    Logger.log('Executed Query: ' + query_string);
    var results = stmt.executeQuery(query_string);
    var metaInfo = results.getMetaData();
    var numCols = metaInfo.getColumnCount();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var active_row = 3;  // recorders start with row 3
    
    // Clear Range
    sheet.getRange(active_row, 1, lastRow ,lastCol).clear();
    
    // Set Column Name
    for (var col = 0; col < numCols; col++) {
      var colName = metaInfo.getColumnName(col + 1)
      sheet.getRange(active_row, col + 1).setValue(colName)
      sheet.getRange(active_row, col + 1).setBackground("#b5d2ff")      
    }
    active_row += 1
    
    // Fetch recorders
    while (results.next()) {
      for (var col = 0; col < numCols; col++) {
        var rowString = results.getString(col + 1);
        sheet.getRange(active_row, col + 1).setValue(rowString)
      }
      active_row += 1;
    }
    
    results.close();
    stmt.close();
    
    Browser.msgBox('Completed!');
    
  } catch(e) {
    Browser.msgBox('Select Error: ' + e);
    Logger.log('Select Error Occured!');
    Logger.log(e);    
  }
}

// Message box
function makeSureMsgBox(msg) {
  var makeSure = Browser.msgBox(msg, Browser.Buttons.OK_CANCEL);
  if (makeSure == 'cancel') {
    Browser.msgBox('Canceled!');
    return false;
  } else {
    return true;
  }
}

function updateData(e) {
  try {
    // Make Sure Msg Box
    var check = makeSureMsgBox("Are you sure to update the data?");
    if (check == false) {
      return;
    }
    
    // Start Update
    Logger.log('Try Update Data');
    var sheet = SpreadsheetApp.getActiveSheet();
    var table_name = sheet.getRange('B2').getValue();
    var conn = getConnection();
    var stmt = conn.createStatement();
    var check_query_string = "SHOW TABLES LIKE '" + table_name + "'";
    var check_result = stmt.executeQuery(check_query_string)
    var metaInfo = check_result.getMetaData();
    Logger.log('check_query_string =>' + check_query_string);
    Logger.log('metainfo =>' + metaInfo.getLastRow);
    if (check_result.first() == false) {
      // Table is not exists
      msgLog = 'There is no table, check your table name please!';
      Browser.msgBox(msgLog);
      Logger.log(msgLog);
      check_result.close();
      return ;
    } 
    check_result.close();

    // Delete All Data
    var delete_query_string = 'DELETE FROM ' + table_name;
    Logger.log('Deleted String: ' + delete_query_string); 
    var delete_result = stmt.executeUpdate(delete_query_string);
    Logger.log('Deleted rows: ' + delete_result);

    // Insert All Data
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var active_row = 3;

    for (var i = active_row; i < lastRow; i++) {
      var data_array = []
      // Check pk
      // var firstCellValue = sheet.getRange(i + 1, 1).getValue();
      // if (firstCellValue == '') {
      //  continue;
      // }
      for (var j = 0; j < lastCol; j++) {
        var cell_data = sheet.getRange(i + 1, j + 1).getValue();
        data_array.push("\"" + cell_data + "\"");
      }
      var values = data_array.join(',');
      var insert_query_string = 'INSERT INTO ' + table_name + ' VALUES ' + '(' + values + ')';
      Logger.log('Inserted String: ' + insert_query_string);  
      var insert_result = stmt.executeUpdate(insert_query_string);
      Logger.log('Inserted : ' + insert_result);  
    }

    stmt.close();
    
    Browser.msgBox('Completed!');
    
  } catch(e) {
    Browser.msgBox('Update Error: ' + e);
    Logger.log('Update Error Occured!');
    Logger.log(e);  
  }
}
