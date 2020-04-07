//Set of Functions that can be attached to a Google Spreadsheet to facilitate a Google Drive Shared Folder to Team Drive Migration
//Author: Jared Garcia
//Attach to a Spreadsheet with three Sheets
//Link to spreadsheet template: https://docs.google.com/spreadsheets/d/1TPPx1XrEcyXRurZWKpy9Qkm0HSpbkFG3FI8q2jzukuE/edit?usp=sharing
//First Sheet needs to be named "Folder Analysis" and havethe following Data in the first row Folder Name, Folder ID, Owner, Parent Folder Name, Parent Folder ID, Gather Folders [NO/YES], Gather Files [NO/YES], 
//Team Drive Duplicate Name,TD Duplicate ID, File Move [NO/YES], File Move Complete
//The Second Sheet should be named "File Analysis" and have the following data  File ID, Owner, Folder Parent, Parent ID, Team Drive Move (Y/N)
//Third Sheet is called "Change Owner Commands". This sheet isn't necessary but is handy when you need to change a file/folders owner using Google Apps Manager commands.
//Version. 0.8 
//This is not a completed script as of yet. It currently gathers folder/files data and copies folders/files to a Team Drive. There is also a deleteFiles function to start over the files data transfer.
//For the most part this can move over a whole shared folder structure using a Time-based trigger on dataAnalysis() until it is finished and then change the trigger to activate migrateData().
//I need to build in Validation however and write some install/delete trigger functions with Menu Add functions to complete as a final deliverable.
function migrateData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Folder Analysis");
  var ssF = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("File Analysis");
  var ssLRow = ss.getLastRow();
  var ssData = ss.getRange(2, 1, ssLRow, 11).getValues();
  findFolders(ssData);
  findFiles(ssF,ssData);
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Folder Analysis");
  var ssF = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("File Analysis");
  var ssLRow = ss.getLastRow();
  var ssData = ss.getRange(2, 1, ssLRow, 11).getValues();
  createFolders(ssData);
  copyFiles(ssF,ssData);
}  

function dataAnalysis(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Folder Analysis");
  var ssF = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("File Analysis");
  var lastRow = ss.getLastRow();
  var lastRowF = ssF.getLastRow();
  var folderCheck = ss.getRange(2, 6, lastRow, 1).getValues();
  var folderResult = checkComplete(folderCheck);
  var fileCheck = ss.getRange(2, 7, lastRow, 1).getValues();
  var fileResult = checkComplete(fileCheck);
  Logger.log(fileResult);
  
  if(folderResult == false) {
    grabFolders(ss, ssF);
  }  
  
  if(fileResult == false) {
    grabFiles(ss, ssF);
  }  
  
}  

//Worker Functions*********************************************************************************
//*************************************************************************************************

//Migrate Data Functions***************************************************************************
//*************************************************************************************************

function findFolders(ss) {
  var iterations = 0;
  for(var i = 0; i < ss.length; i++) {
    var present;
    if(ss[i][8] == "N/A") {
      var parentName = ss[i][3];
      var parentFolder = getParentTDFolder(ss[i][4],ss);
      Logger.log(parentName);
      if(parentFolder != "Finish") {
        var folders = parentFolder.getFolders();
        while(folders.hasNext()) {
          var folder = folders.next();
          if(ss[i][7] == folder.getName()) {
            var number = i+2;
            var cell = "I"+number;
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Folder Analysis").getRange(cell).setValue(folder.getId());
          }  
        }  
      }
      if(parentName != ss[i+1][3]){
        Logger.log(iterations);
        iterations++;
        if(iterations > 50) {
          i = 9999999999999999999999999;
        }
      }
    }
  }
}  

function createFolders(ss) {
  var iterations = 0;
  for(var i = 0; i < ss.length; i++) {
    if(ss[i][8] == "N/A") {
      var parentName = ss[i][3];
      var parentFolder = getParentSharedFolder(ss[i][4],ss);
      parentFolder.createFolder(ss[i][7]);
      if(parentName != ss[i+1][3]) {
        iterations++;
        if(iterations > 50) {
          i = 9999999999999999999999999;
        }  
      }
    }
    
  }  
  
}  

function getParentSharedFolder(id, ss) {
  for(var i = 0; i < ss.length; i++) {
    if(ss[i][1] == id) {
      Logger.log(ss[i][8]);
      var parent = DriveApp.getFolderById(ss[i][8]);
      return parent;
    }  
  }
}

function getParentTDFolder(id, ss) {
  //Logger.log(id);
  for(var i = 0; i < ss.length; i++) {
    if(ss[i][1] == id) {
      try {
        var parent = DriveApp.getFolderById(ss[i][8]);
      } catch(e) {
        //Logger.log("That Parent Folder hasn't been made yet");
        var parent = "Finish";
      }  
      i = 999999999999999;
    }  
  }
  return parent;
} 

function copyFiles(ssF,ss) {
  var iterations = 0;
  var lastRow = ssF.getLastRow();
  var data = ssF.getRange(2, 1, lastRow, 6).getValues();
  for(var i = 0; i < data.length; i++) {
    if(data[i][5] == "NO") {
      Logger.log("This worked!");
      var parent = getParentTDFolder(data[i][4],ss);
      DriveApp.getFileById(data[i][1]).makeCopy(data[i][0], parent);
      if(data[i][3] != data[i+1][3]) {
        iterations++;
        if(iterations > 50) {
          i = 9999999999999999999999999;
        }
      }
    }  
  }  
}

function findFiles(ssF,ss) {
  var lastRow = ssF.getLastRow();
  var data = ssF.getRange(2, 1, lastRow, 6).getValues();
  var iterations = 0;
  for(var i = 0; i < data.length; i++) {
    var present;
    if(data[i][5] == "NO") {
      var parentFolder = getParentTDFolder(data[i][4],ss);
      Logger.log(parentFolder);
      if(parentFolder != "Finish") {
        var files = parentFolder.getFiles();
        while(files.hasNext()) {
          var file = files.next();
          if(data[i][0] == file.getName()) {
            var number = i+2;
            var cell = "F"+number;
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("File Analysis").getRange(cell).setValue("YES");
          }  
        }
        if(data[i][3] != data[i+1][3]){
          iterations++;
          if(iterations > 50) {
            i = 9999999999999999999999999;
          }
        }  
      }  
    }
  }
}


//Worker Analysis Functions************************************************************************
//Grabs Folder data from the sheet to pass to the writeJustFolders() function
function grabFolders(ss, ssF) {
  
  var lastRow = ss.getLastRow();
  var lastRowF = ssF.getLastRow();
  var iterativeRange = ss.getRange(1, 1, lastRow, 6).getValues();
  var placeHolder = 0;
  Logger.log(iterativeRange);
  for(var i = 0; i < iterativeRange.length; i++) {
    
    if(iterativeRange[i][5] == "NO") {
      var writeArray = [];
      Logger.log(placeHolder);
      if(placeHolder < 75) {
        placeHolder++;
        writeArray = writeJustFolders(iterativeRange[i][1], 1, iterativeRange[i][0],writeArray);
        ss.getRange(i+1, 6).setValue("YES");
        Logger.log(writeArray);
        if(writeArray.length != 0){
          ss.getRange(lastRow+1, 1, writeArray.length, 8).setValues(writeArray);
          lastRow = lastRow + writeArray.length;
        }
      }  
    }  
    
  }
  
}

//Uses data from grabFolders() to find and write data on Folders to the spreadsheet
function writeJustFolders(folderId, row, parentName, writeArray) {  
  
  var writeFolder = DriveApp.getFolderById(folderId);
  var folders = writeFolder.getFolders();
  while (folders.hasNext()) {
    var writeThis = [];
    var folder = folders.next();
    var owner = folder.getOwner();
    writeThis.push(folder.getName());
    writeThis.push(folder.getId());
    writeThis.push(owner.getEmail());
    writeThis.push(parentName);
    writeThis.push(folderId);
    writeThis.push("NO");
    writeThis.push("NO");
    writeThis.push(folder.getName());
    writeArray.push(writeThis);
  }  
  
  return writeArray;
}

//Grabs Folder data from the sheet to pass to the writeJustFiles() function
function grabFiles(ss, ssF) {
  
  var lastRow = ss.getLastRow();
  var lastRowF = ssF.getLastRow();
  var iterativeRange = ss.getRange(1, 1, lastRow, 7).getValues();
  var placeHolder = 0;
  for(var i = 0; i < iterativeRange.length; i++) {
    Logger.log(iterativeRange[i][6]);
    if(iterativeRange[i][6] == "NO") {
      var writeArray = [];
      Logger.log(placeHolder);
      if(placeHolder < 75) {
        placeHolder++;
        writeArray = writeJustFiles(iterativeRange[i][1], 1, ss,writeArray);
        ss.getRange(i+1, 7).setValue("YES");
        Logger.log(lastRow);
        if(writeArray.length != 0){
          ssF.getRange(lastRowF+1, 1, writeArray.length, 6).setValues(writeArray);
          lastRowF = lastRowF + writeArray.length;
        }
      }  
    }  
    
  }
  
  
  
}

//Uses data from grabFiles() to find and write data on the files contained in Folders to the spreadsheet
function writeJustFiles(folderId, row, writeSheet,writeArray) {
  
  
  //Logger.log(folderId);
  var writeFolder = DriveApp.getFolderById(folderId);
  var files = writeFolder.getFiles();
  while (files.hasNext()) {
   var writeThis = [];
   var file = files.next();
   var owner = file.getOwner();
   writeThis.push(file.getName());
   writeThis.push(file.getId());
   writeThis.push(owner.getEmail());
   writeThis.push(writeFolder.getName());
   writeThis.push(writeFolder.getId());
   writeThis.push("NO"); 
   writeArray.push(writeThis);
  }  
  
  return writeArray;
} 

//Builds Change Owner Commands
function buildGAMcommandOwner(sheet) {
  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var writeThis = [];
  for(var i = 0; i < values.length; i++)  {
    if(values[i][2] != "user@example.com") {
      var pushThis = ["gam user "+ values[i][2] + " update drivefileacl " + values[i][1] + " user@example.com role owner transferownership true"];
      writeThis.push(pushThis);
    }  
    
  }
  var writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Change Owner Commands");
  var lastRow = writeSheet.getLastRow() + 1;
  if(writeThis.length != 0) {
    writeSheet.getRange(lastRow,1,writeThis.length).setValues(writeThis); 
  }
}

function checkComplete(array) {
  
  for(var i = 0; i < array.length; i++) {
    if(array[i][0] == "NO") {
      var finished = false;
      return finished;
    }  
  }

  return finished;  
  
}

//Use this function if you need to rollback file moves to start over from a clean slate**********************************
//***********************************************************************************************************************
function deleteFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Folder Analysis");
  var ssF = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("File Analysis");
  var ssLRow = ss.getLastRow();
  var ssFLRow = ssF.getLastRow();
  var folderData = ss.getRange(1621,9, ssLRow, 1).getValues();
  for(var i = 0; i < folderData.length; i++) {
    var folder = DriveApp.getFolderById(folderData[i][0]);
    Logger.log(folder.getName());
    var files = folder.getFiles();
    while(files.hasNext()) {
      files.next().setTrashed(true);
    }  
  }  
  
}  

//Build/Delete Triggers Functions*************************************************************************
//********************************************************************************************************
function analysisTimer(){
  ScriptApp.newTrigger('dataAnalysis')
     .timeBased()
     .everyMinutes(15)
     .create();
  
  
}  

function deleteTimer(){
  

}

function deleteTrigger(){
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }  
  
}  
//Menu Functions******************************************************************************************
//********************************************************************************************************
function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('Migration Tools')
       .addItem('Start Analysis', 'analysisTimer')
       .addItem('Stop Analysis', 'deleteTimer')
       .addItem('Migrate Data', 'migrateData')
       .addItem('Stop Migration', 'deleteTrigger')
       .addToUi();
} 
 
