/*
 *   KINGS CROSS
 *
 *   This script takes information from a Firebase server
 *   and then pushes the data to the spreadsheet.
 */

function dump2(){
  
  // gets current users email
  var email = Session.getActiveUser().getEmail();
 
  // does not run if the current user is not me or Brother Moore
  if(email != "ericjulander@gmail.com" && email != "mooreco.byui@gmail.com"){
    alert(email+" is not allowed to use this button");
    return;
  }
 
  
  // variables for Firebase data extraction
  
  var firebaseUrl = "https://mmap.firebaseio.com/";
  var secret = "rs4XlisWL1eJNTdWkuUx9LrYdH8b8xilyerrJz5B";
  var base = FirebaseApp.getDatabaseByUrl(firebaseUrl, secret);
  
  var data = base.getData();
  var station = SpreadsheetApp.getActive().getSheetByName("Kings Cross");
  var teams = station.getRange(2, 1,station.getLastRow(),1).getValues();

  var courses = [];
  var currentRow = 3;
  
  
  // gets target destination for Firebase data
  var sheet = SpreadsheetApp.getActive().getSheetByName("Spring 2016");
  if(sheet.getLastRow() > 3)    
  sheet.deleteRows(3, sheet.getLastRow()-3);
 
 // parse Firebase JSON data into variables to insert into spreadsheet
 var dataObject = new Array(sheet.getLastRow()-3); 
  for(var t in teams){
    var current = teams[t][0];
    
    if( current != "" ){
      
      var team = parseName(current);
      for(var i in data[current])
      {
        var courseName = i;
        
       
        var level = data[current][i];
        var dashboard = '=HYPERLINK("'+level['details'].dashboard+'",">X<")' ;
        var status = level['details'].status;
        var week = level['details'].week;
        
        
        var cdEmail = level["dir"]["cd"].email;
        var cdName = level["dir"]["cd"].name;
        
        var clEmail = level["dir"]["cl"].email;
        var clName = level["dir"]["cl"].name;
        
        
        var slName = level["dir"]["sl"].name;
        
        var il2 = '=HYPERLINK("'+level["links"].il2+'",">X<")' ;
        var il3 = '=HYPERLINK("'+level["links"].il3+'",">X<")' ;
        
        
        var progress = level["progress"].phase;
        
        var things = [courseName, dashboard, status, progress,"=INDIRECT(ADDRESS(MATCH(D"+currentRow+",Variables!$C$1:$C$33,0),4,1,1,\"Variables\"))",slName, team,week,clName,clEmail, cdEmail, cdName, il2, il3];
        dataObject[currentRow-3] = things;
        
        
        
        currentRow++;
          
      }
    }
  }
  
  // inserts new row
  sheet.getRange(3, 1, dataObject.length, 14).setValues(dataObject);
  sortByWeek();

}


/*
 * This sorts the spreadsheet according to the week data.
 */
function sortByWeek(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Spring 2016");
  sheet.getRange(3,1,sheet.getLastRow()-2, sheet.getLastColumn()).sort(8);
}

/*
 * Changes the data from the format "HouseNumber" to HT#
 * Example: "GryffindorOne" returns "GT01" 
 */
function parseName(name){
  var words = splitWord(name);
  words[1] = wordToNumber(words[1]);
  return words[0][0] + "T "+words[1];
}

/*
 * Splits word by captial letters
 */
function splitWord(word){
  word = word.trim();
  var words = ["",""];
  var location = 0;
  var caps = 0;
  var len = word.length;

  for(var i = 0; i < len; i++){
      if(word[i] == word[i].toUpperCase()){
        if(caps >= 1){
          location = i;
      
          break;
        }
        else{
          caps++;
        }
      } 
  }
  
  words[0] = word.substr(0, location);
  words[1] = word.substr(location);
  
  return words;
  
}

/*
 * Converts a word to number
 * Example: "One" returns "1"
 */
function wordToNumber(word){
  
  var number  = "00";
  switch(word.toLowerCase()){
    case "one":
      number =  "01";
      break;
     case "two":
      number =  "02";
      break;
      
      case "three":
      number =  "03";
      break;
      
      case "four":
      number =  "04";
      break;
      
      case "five":
      number =  "05";
      break;
      
      case "six":
      number =  "06";
      break;
      
      case "seven":
      number =  "07";
      break;
      
      case "eight":
      number =  "08";
      break;
      
      case "nine":
      number =  "09";
      break;
      
      case "ten":
      number =  "10";
      break;
      
      
  }
  
  return number;

}


/*
 * Makes a popup message on the spreadsheet
 */
function alert(string){
  SpreadsheetApp.getUi().alert(string);
}

