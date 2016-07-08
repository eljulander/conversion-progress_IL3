/*
 *   DASHBOARD GEN
 *  
 *   This takes information from the active row of the spreadsheet and creates a new dashboard.
 *   After the dashboard has been created, the second column of the selected row will be filled
 *   with a link to the dashboard.  
 */




// Global variables for the spreadsheets
var sheet = SpreadsheetApp.getActive();
var conversionSheet = sheet.getSheetByName("Spring 2016");
var template = "18wHsI2mY4VBmQoDas3mUzOyynEWNwyKvUMmeAPPKwkE";
var currentRow = conversionSheet.getActiveCell().getRow();
var tempFolder = "0B-XYjE5fqvtQNG5NSGlhUzU5RzA";


/*
 * Creates dashboard menu options in spreadsheet
 */
function onOpen(e){
  
  // creates menu
  var menu = SpreadsheetApp.getUi();
  
  menu.createMenu("Dashboard")
  .addItem("Create Dashboard", "generateDashboard")
  .addItem("Update Dashboard", "placeHolder")
  .addToUi();
 
}


/*
 * Construction Redirect for the Menu
 */
function placeHolder(){
  var menu = SpreadsheetApp.getUi();
  menu.alert("This function is currently unavaliable");
}

/*
 * Creates Dashboard
 */
function generateDashboard(){

 // creates a new dashboard in the specified folder
  var sheet = SpreadsheetApp.getActive();
  var name  = SpreadsheetApp.getActiveSheet().getRange(currentRow, 1).getValue();
  var folder =  DriveApp.getFolderById("0B-XYjE5fqvtQNG5NSGlhUzU5RzA").createFolder(name);
  var newSheet = DriveApp.getFileById(template).makeCopy(folder);
  newSheet.setName(name + " Dashboard");
 
// stores the new dashboard into a variable
  var cs = SpreadsheetApp.openById(newSheet.getId());
  var team = sheet.getActiveSheet().getRange(currentRow, 7).getValue();
      
      
      
   
      
  
  
  
  // transfer of variables
  

  // status
  cs.getSheets()[0].getRange(3, 4).setValue(sheet.getActiveSheet().getRange(currentRow, 3).getValue());
  
  // progress
  cs.getSheets()[0].getRange(3, 5).setValue(sheet.getActiveSheet().getRange(currentRow, 4).getValue());
  
  // student lead
  cs.getSheets()[0].getRange(3, 7).setValue(sheet.getActiveSheet().getRange(currentRow, 6).getValue());
  
  // team lead
  cs.getSheets()[0].getRange(3, 8).setValue(findLead(team));
  
  // course lead
  cs.getSheets()[0].getRange(6, 2).setValue(sheet.getActiveSheet().getRange(currentRow, 9).getValue());
   
  // course lead email
  cs.getSheets()[0].getRange(6, 5).setValue(sheet.getActiveSheet().getRange(currentRow, 10).getValue());
  
  // curriculum Designer
  cs.getSheets()[0].getRange(6, 6).setValue(sheet.getActiveSheet().getRange(currentRow, 12).getValue());
  
  // curriculum Designer email
  cs.getSheets()[0].getRange(6, 7).setValue(sheet.getActiveSheet().getRange(currentRow, 13).getValue());
  

  // week
  cs.getSheets()[0].getRange(6, 7).setValue(sheet.getActiveSheet().getRange(currentRow, 8).getValue());
  
  // team
  cs.getSheets()[0].getRange(6, 8).setValue(team);
  
  
  
  

 
  // creates variable sheet in the new dashboard
  
  var varSheet = cs.insertSheet("vars");
  varSheet.protect().addEditor("ericjulander@gmail.com");
  varSheet.protect().addEditor("mooreco.byui@gmail.com");
  
  // index 
  varSheet.getRange(1, 1).setValue("Index:");
  varSheet.getRange(1, 2).setValue(currentRow);
  
  // creator
  varSheet.getRange(2, 1).setValue("Spreadsheet Created By:");
  varSheet.getRange(2, 2).setValue( Session.getActiveUser().getEmail());
  
  // mother 
  varSheet.getRange(3, 1).setValue("Mother:");
  varSheet.getRange(3, 2).setValue( SpreadsheetApp.getActive().getId());
  
  // mother Sheet
  varSheet.getRange(4, 1).setValue("Mother Sheet:");
  varSheet.getRange(4, 2).setValue( SpreadsheetApp.getActive().getActiveSheet().getName());
  
  varSheet.hideSheet();
  
  // generate link
  var url = newSheet.getUrl();
  sheet.getSheetByName("Spring 2016").getRange(currentRow, 2).setValue('=HYPERLINK("'+url+'",">X<")');
      
} 


/*
 * Finds home row
 */
function getRow(name , sheet){
  var index = -1;
  
  for(var i in sheet){

    if(sheet[i][0] == name){
      Logger.log(sheet[i][0]+" "+i);
      index = parseInt(i);
    }
    
  }
  
  return index+1;
}



/*
 * This finds the correct data location from the mother sheet
 */
function indexer(name,sheet){
  var index = -1;
  for(var a in sheet[0]){
    if(sheet[0][a] == name){
      index = a;
      break;
    }
  }
  return parseInt(index);
  
}

/*
 * This finds the correct data location on the dashboard
 */
function indexer2(name){
  var index = -2;
  var row = -2;
  var sheet = SpreadsheetApp.getActive().getSheets()[0].getDataRange().getValues();
  for(var b in sheet){
    for(var a in sheet[0]){
      if(sheet[b][a] == name){
        index = a;
        row = b;
        break;
      }
    }
  }
  return [parseInt(index)+1, parseInt(row)+1];
}




