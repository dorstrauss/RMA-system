//a function that retrive the data of the requested RMA
function rmaDataRetrive() 
{
  var buttonPressed = SpreadsheetApp.getUi().alert("If you retrive the data of this RMA number all the changes you made in the bottom table will be deleted. Are you sure you want to continue? ",SpreadsheetApp.getUi().ButtonSet.YES_NO);

  if (buttonPressed == SpreadsheetApp.getUi().Button.NO || buttonPressed == SpreadsheetApp.getUi().Button.CLOSE) //if the user decided not to continue
  {
    return;
  }
  
    cleanCells(); //cleaning all the cells befor entering new data

    var thisSheet = SpreadsheetApp.getActiveSheet();
    thisSheet.getRange("RMA").setNumberFormat("@"); //seting the rma 
    var requstedRmaId = thisSheet.getRange("RMA").getValue(); //gets the rma the user want to retrive data about

    //opening the database sheet where the customers form values are stored
    var rmaDatabaseSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1B3PIk0T0F4txYTE9daP1pVCdhH2t_zVGd6Uz_SLuZi4/edit?usp=sharing");
    var rmaDatabaseSheet = rmaDatabaseSpreadsheet.getSheetByName("CUSTOMERS RMA");

    var rmaExist = false; //setting a boolean variable so if the rma does not exist in the database is will notify
    rmaDatabaseSheet.getRange("A:A").setNumberFormat('@'); //Seting the cells format to text so when sorting the column numbers and text will be ordered in the right order and the binary search will work
    var filterExist = rmaDatabaseSheet.getFilter();
    if (filterExist == null) //if there is no filter in the sheet
    {
      rmaDatabaseSheet.getRange("A:U").createFilter();
    }
    rmaDatabaseSheet.getFilter().sort(1, true); //sort the customers sheet rma column from small to large
    
    //preforming a binary search on the values in the customres database to retrive the requested rma data.
    var rangeEnd = rmaDatabaseSheet.getRange("A:A").getLastRow(); //getting the last row of data in the sheet
    if (rmaDatabaseSheet.getRange(rangeEnd,1).getValue() == "")
    {
      rmaDatabaseSheet.getRange(rangeEnd,1).activate();
      rmaDatabaseSheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
      rangeEnd = rmaDatabaseSheet.getActiveCell().getRowIndex();
    }
    var rangeStart = 2;
    var isFirstRowCopyied = false;
    var copyRow = 14;
    while (rangeStart <= rangeEnd)
    {
      var rangeMiddel = Math.floor((rangeStart + rangeEnd)/2); //calculating the middle of the range

      if (requstedRmaId == rmaDatabaseSheet.getRange(rangeMiddel,1).getValue()) //if the middle of the search range is equal to the requsted rma id
      {
        while (requstedRmaId == rmaDatabaseSheet.getRange(rangeMiddel - 1,1).getValue()) //going to the first row of the requsted rma.
          {
            rangeMiddel = rangeMiddel - 1;
          }
        while (requstedRmaId == rmaDatabaseSheet.getRange(rangeMiddel,1).getValue())
        {
          if (isFirstRowCopyied == false) //if it is the first row we copy the general detail of the rma
          {
            rmaExist = true;
            thisSheet.getRange("Company").setValue(rmaDatabaseSheet.getRange(rangeMiddel,9).getValue());
            thisSheet.getRange("Contact").setValue(rmaDatabaseSheet.getRange(rangeMiddel,10).getValue());
            thisSheet.getRange("Address").setValue(rmaDatabaseSheet.getRange(rangeMiddel,11).getValue());
            thisSheet.getRange("phoneNumber").setValue(rmaDatabaseSheet.getRange(rangeMiddel,12).getValue());
            thisSheet.getRange("City").setValue(rmaDatabaseSheet.getRange(rangeMiddel,13).getValue());
            thisSheet.getRange("State").setValue(rmaDatabaseSheet.getRange(rangeMiddel,14).getValue());
            thisSheet.getRange("Email").setValue(rmaDatabaseSheet.getRange(rangeMiddel,16).getValue());
            thisSheet.getRange("OrderDate").setValue(rmaDatabaseSheet.getRange(rangeMiddel,18).getValue());
            isFirstRowCopyied = true;
          }

          //copying this row data to the sheet
          thisSheet.getRange(copyRow,2).setValue(rmaDatabaseSheet.getRange(rangeMiddel,3).getValue());
          thisSheet.getRange(copyRow,2).setHorizontalAlignment('center');
          thisSheet.getRange(copyRow,3).setValue(rmaDatabaseSheet.getRange(rangeMiddel,4).getValue());
          thisSheet.getRange(copyRow,3).setHorizontalAlignment('center');
          thisSheet.getRange(copyRow,4).setValue(rmaDatabaseSheet.getRange(rangeMiddel,5).getValue());
          thisSheet.getRange(copyRow,4).setHorizontalAlignment('center');
          thisSheet.getRange(copyRow,5).setValue(rmaDatabaseSheet.getRange(rangeMiddel,6).getValue());
          thisSheet.getRange(copyRow,5).setHorizontalAlignment('center');
          thisSheet.getRange(copyRow,6).setValue(rmaDatabaseSheet.getRange(rangeMiddel,7).getValue());
          thisSheet.getRange(copyRow,6).setHorizontalAlignment('center');
          thisSheet.getRange(copyRow,7).setValue(rmaDatabaseSheet.getRange(rangeMiddel,8).getValue());
          thisSheet.getRange(copyRow,7).setHorizontalAlignment('center');
          thisSheet.getRange(copyRow,8).setValue(rmaDatabaseSheet.getRange(rangeMiddel,19).getValue());
          thisSheet.getRange(copyRow,8).setHorizontalAlignment('center');
          
          copyRow++;
          rangeMiddel++;
        }
        break; //ending the search because we found the rma
      }
      else if (requstedRmaId < rmaDatabaseSheet.getRange(rangeMiddel,1).getValue()) //if the requested rma is smaller than the current middle search.
      {
        rangeEnd = rangeMiddel - 1;
      }
      else //if the requested rma is bigger than the current middle search
      {
        rangeStart = rangeMiddel + 1;
      }
    }

    rmaDatabaseSheet.getFilter().sort(2,true); //canceling any column filter before pasteing the new values

    if (rmaExist == false) //if the requsted RMA ID does not exist in the customers database.
    {
      SpreadsheetApp.getUi().alert("The RMA you looking for does not exist in the database");
      return;
    }
    //continue to check if the fixtures allready been fixed in the CM database, and retreive there data.
    
      var rmaCmDatabaseSheet = rmaDatabaseSpreadsheet.getSheetByName("CM RMA");

      copyRow = 14;


      //preforming a binary search to find the RMA ID in the CM database
      rmaCmDatabaseSheet.getRange("A:A").setNumberFormat('@'); //Seting the cells format to text so when sorting the column numbers and text will be ordered in the right order and the binary search will work
      var filter = rmaCmDatabaseSheet.getFilter();
      if (filter == null)
        {rmaCmDatabaseSheet.getRange("A:BP").createFilter();}
      rmaCmDatabaseSheet.getFilter().sort(1, true); //sorting the RMA ID column from small to large

      rangeStart = 2;
      rangeEnd = rmaCmDatabaseSheet.getRange("A:A").getLastRow();
      if (rmaCmDatabaseSheet.getRange(rangeEnd,1).getValue() == "") //getting the last row of data
      {
        rmaCmDatabaseSheet.getRange(rangeEnd,1).activate();
      rmaCmDatabaseSheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
      rangeEnd = rmaCmDatabaseSheet.getActiveCell().getRowIndex();
      }

      while (rangeStart <= rangeEnd)//continue the binary search until we coverd all the range
      {
        var rangeMiddel = Math.floor((rangeStart + rangeEnd)/2); //calculating the middle of the range.
        var rangeMiddelValue = rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue();

        if (requstedRmaId == rangeMiddelValue) //if this row match the requested rma
        {
          while (requstedRmaId == rmaCmDatabaseSheet.getRange(rangeMiddel-1,1).getValue()) //going to the first requested rma row
          { 
            rangeMiddel = rangeMiddel - 1;
          }

          while (requstedRmaId == rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue()) //going through all the rows with the requested rma
          { 
            copyRow = 14;
            while (thisSheet.getRange(copyRow,5).getValue() != "") //we didn't go through all the rows in the form
            {
              if (rmaCmDatabaseSheet.getRange(rangeMiddel,13).getValue() == thisSheet.getRange(copyRow,5).getValue()) //if the luminer in tha data base as the same serial number as the one in the technican form
              {
                for (var column = 17;column <= 30;column++) //copying all the row data from the databse
                {
                  thisSheet.getRange(copyRow,column - 8).setValue(rmaCmDatabaseSheet.getRange(rangeMiddel,column).getValue());
                }

                for (var column = 32;column <= 48;column++)
                {
                  thisSheet.getRange(copyRow,column - 8).setValue(rmaCmDatabaseSheet.getRange(rangeMiddel,column).getValue());
                }
                copyRow++;
                break;
              }
              else
                copyRow++;
            }
            rangeMiddel++;
          }
          break;
        }
        else if (requstedRmaId <  rangeMiddelValue) //if the requested rma is smaller than the current middle.
        {
          rangeEnd = rangeMiddel - 1;
        }
        else if (requstedRmaId >  rangeMiddelValue) //if the requested rma is bigger than the current middle.
        {
          rangeStart= rangeMiddel + 1;
        }
      }
    

}












//a function that cleans all the cells
function cleanCells()
{
  var thisSheet = SpreadsheetApp.getActiveSheet();
  thisSheet.getRange("Company").setValue("");
  thisSheet.getRange("Contact").setValue("");
  thisSheet.getRange("Address").setValue("");
  thisSheet.getRange("City").setValue("");
  thisSheet.getRange("State").setValue("");
  thisSheet.getRange("Email").setValue("");
  thisSheet.getRange("OrderDate").setValue("");
  thisSheet.getRange("phoneNumber").setValue("");
  thisSheet.getRange("rmaDecision").setValue("");
  thisSheet.getRange("Capa1").setValue("");
  thisSheet.getRange("Capa2").setValue("");
  thisSheet.getRange("Capa3").setValue("");
  thisSheet.getRange("Capa4").setValue("");
  thisSheet.getRange("Capa5").setValue("");
  thisSheet.getRange("technicianName").setValue("");
  thisSheet.getRange("fixCountry").setValue("");
  thisSheet.getRange("technicianCode").setValue("");

  var currentRow = 14;
  while (thisSheet.getRange(currentRow,2).isBlank() == false)
  {
    thisSheet.getRange(currentRow,2).setValue("");
    thisSheet.getRange(currentRow,3).setValue("");
    thisSheet.getRange(currentRow,4).setValue("");
    thisSheet.getRange(currentRow,5).setValue("");
    thisSheet.getRange(currentRow,6).setValue("");
    thisSheet.getRange(currentRow,7).setValue("");
    thisSheet.getRange(currentRow,8).setValue("");
    thisSheet.getRange(currentRow,9).setValue("");
    thisSheet.getRange(currentRow,10).setValue("");
    thisSheet.getRange(currentRow,11).setValue("FALSE");
    thisSheet.getRange(currentRow,12).setValue("FALSE");
    thisSheet.getRange(currentRow,13).setValue("FALSE");
    thisSheet.getRange(currentRow,14).setValue("FALSE");
    thisSheet.getRange(currentRow,15).setValue("FALSE");
    thisSheet.getRange(currentRow,16).setValue("FALSE");
    thisSheet.getRange(currentRow,17).setValue("FALSE");
    thisSheet.getRange(currentRow,18).setValue("FALSE");
    thisSheet.getRange(currentRow,19).setValue("FALSE");
    thisSheet.getRange(currentRow,20).setValue("");
    thisSheet.getRange(currentRow,21).setValue("");
    thisSheet.getRange(currentRow,22).setValue("");
    thisSheet.getRange(currentRow,24).setValue("");
    thisSheet.getRange(currentRow,25).setValue("FALSE");
    thisSheet.getRange(currentRow,26).setValue("");
    thisSheet.getRange(currentRow,27).setValue("FALSE");
    thisSheet.getRange(currentRow,28).setValue("");
    thisSheet.getRange(currentRow,29).setValue("FALSE");
    thisSheet.getRange(currentRow,30).setValue("");
    thisSheet.getRange(currentRow,31).setValue("FALSE");
    thisSheet.getRange(currentRow,32).setValue("");
    thisSheet.getRange(currentRow,33).setValue("");
    thisSheet.getRange(currentRow,34).setValue("FALSE");
    thisSheet.getRange(currentRow,35).setValue("");
    thisSheet.getRange(currentRow,36).setValue("");
    thisSheet.getRange(currentRow,37).setValue("FALSE");
    thisSheet.getRange(currentRow,38).setValue("");
    thisSheet.getRange(currentRow,39).setValue("");
    thisSheet.getRange(currentRow,40).setValue("FALSE");

    currentRow++;
  }

}











//a function that beeing called after the technician finished the RMA form and send the new data to the database.
function rmaCmUpdate()
{
  var thisSheet = SpreadsheetApp.getActiveSheet();

  var emptyCriticalCells = false;
  var currentRow = 14;
  //checking if the technician filled all the necessary fileds in the form
  while (thisSheet.getRange(currentRow,2).isBlank() == false)
  {
    if (thisSheet.getRange(currentRow,9).getValue() == "" || thisSheet.getRange(currentRow,11,1,3).isChecked() == false || thisSheet.getRange(currentRow,11,1,3).isChecked() == null || thisSheet.getRange(currentRow,15,1,5).isChecked() == false || thisSheet.getRange(currentRow,15,1,5).isChecked() == null || thisSheet.getRange(currentRow,35).isChecked() == false || thisSheet.getRange("rmaDecision").getValue() == "" || thisSheet.getRange("technicianName").getValue() == "")  //if one or more of the critical feilds are empty
    {
      emptyCriticalCells = true;
      break;
    }
    currentRow++;
  }

  if ( emptyCriticalCells == true) // if one or more of the critical feilds are empty a messaga appear
  {
    var buttonPressed = SpreadsheetApp.getUi().alert('One or more of this fields are not filled (Final RMA decision:,Technician name:,Visual Inspection:,Power ON Check(AC):,Leds Check:,Driver Test:,Dimmer Card Functionally:,Script Test:,Power (W) Test;,Wiring connectivity:,Sealing:).' + ".\n" + 'All those fields should be filled, are you sure you want to continue?',SpreadsheetApp.getUi().ButtonSet.YES_NO);
  
  if (buttonPressed == SpreadsheetApp.getUi().Button.NO || buttonPressed == SpreadsheetApp.getUi().Button.CLOSE) //if the user preesed NO or the 'X' button it will stop the code.
        {
          return;
        }
  }


// //the electronic signnature functions
// function signatureForm() {
//   var html = HtmlService.createHtmlOutputFromFile('signature').setWidth(400).setHeight(300);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Please sign below to complete the submit');
// }


// function saveImage(bytes){
//   var thisSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var rmaNumber = thisSheet.getRange("RMA").getValue();
//   var technicianName = thisSheet.getRange("technicianName").getValue();
//   var formNumber = thisSheet.getRange("formNumber").getValue();
//   var bytes = bytes.split(",");
//   var blob = Utilities.newBlob(Utilities.base64Decode(bytes[1]), 'image/png');
//   blob.setName(technicianName + " sign -RMA:" + rmaNumber + "-form:" + formNumber);
//   var newSignatureId = DriveApp.getFolderById("1jEC8R7tK_XTTRcjOboqq-MqW6Pkz_CE2").createFile(blob).getId();
//   var newSignatureLink = DriveApp.getFileById(newSignatureId).getUrl();
//   databaseUpdate(newSignatureLink);
// }

  var thisSheet = SpreadsheetApp.getActiveSheet();

  if (thisSheet.getRange("technicianName").getValue() == "" || thisSheet.getRange("technicianCode").getValue() == "") //if the technician does not fill it's name or code
  {
    SpreadsheetApp.getUi().alert("You have to fill the 'Technician name' and 'Technician unique code' fields in order to submit the data");
    return; 
  }

  //opening the database sheet where the CM form values are stored
  var rmaDatabaseSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1B3PIk0T0F4txYTE9daP1pVCdhH2t_zVGd6Uz_SLuZi4/edit?usp=sharing");
  var rmaCmDatabaseSheet = rmaDatabaseSpreadsheet.getSheetByName("CM RMA");
  var rmaSummarySheet = rmaDatabaseSpreadsheet.getSheetByName("RMA Summary");

  var fixturesQty = 0; //a variable to sum the number of fixtures in this RMA
  var fixturesSerialNumbers = ""; //a variable to fill all the fixtures S.N
  var replacedPartsSerialNumbers = [,,,,,,,,,,,,,,,,,,,,,,,,];
  var replacedPartsArrayLastIndex = 0; 

  //getting the customer details from the form
  var rmaNumber = thisSheet.getRange("RMA").getValue();
  var company = thisSheet.getRange("Company").getValue();
  var contact = thisSheet.getRange("Contact").getValue();
  var address = thisSheet.getRange("Address").getValue();
  var city = thisSheet.getRange("City").getValue();
  var state = thisSheet.getRange("State").getValue();
  var email = thisSheet.getRange("Email").getValue();
  var phoneNumber = thisSheet.getRange("phoneNumber").getValue();
  var orderDate = thisSheet.getRange("OrderDate").getValue();
  var rmaDecision = thisSheet.getRange("rmaDecision").getValue();
  var Capa1 = thisSheet.getRange("Capa1").getValue();
  var Capa2 = thisSheet.getRange("Capa2").getValue();
  var Capa3 = thisSheet.getRange("Capa3").getValue();
  var Capa4 = thisSheet.getRange("Capa4").getValue();
  var Capa5 = thisSheet.getRange("Capa5").getValue();
  var technicianName = thisSheet.getRange("technicianName").getValue();
  var technicianCode = thisSheet.getRange("technicianCode").getValue();
  var fixCountry = thisSheet.getRange("fixCountry").getValue();

  //if there is no filter in the sheet, we creat a filter
  var filterExist = rmaCmDatabaseSheet.getFilter();
  if (filterExist == null)
  {
    rmaCmDatabaseSheet.getRange("A:BL").createFilter();
  }

  //checking if this RMA number already exist in the CM database.
  var rmaExist = false;
  //preforming a binary search to search the RMA ID
  rmaCmDatabaseSheet.getRange("A:A").setNumberFormat('@'); //Seting the cells format to text so when sorting the column numbers and text will be ordered in the right order and the binary search will work
  rmaCmDatabaseSheet.getFilter().sort(1, true); //sorting the RMA ID column from small to large
  rangeStart = 2;
  rangeEnd = rmaCmDatabaseSheet.getRange("A:A").getLastRow();
  if (rmaCmDatabaseSheet.getRange(rangeEnd,1).getValue() == "") //getting the last row of data
  {
    rmaCmDatabaseSheet.getRange(rangeEnd,1).activate();
    rmaCmDatabaseSheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
    rangeEnd = rmaCmDatabaseSheet.getActiveCell().getRowIndex();
  }

  while (rangeStart <= rangeEnd)//continue the binary search until we coverd all the range
    {
        var rangeMiddel = Math.floor((rangeStart + rangeEnd)/2); //calculating the middle of the range.
        if (rmaNumber == rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue()) //if this row match the requested rma
        {
          rmaExist = true;
          buttonPressed = SpreadsheetApp.getUi().alert('There are already rows with this RMA ID filled by technician in the database, are you want to update them?',SpreadsheetApp.getUi().ButtonSet.YES_NO);
  
          if (buttonPressed == SpreadsheetApp.getUi().Button.NO || buttonPressed == SpreadsheetApp.getUi().Button.CLOSE) //if the user preesed NO or the 'X' button it will stop the code.
          {
          return;
          }
  
          while (rmaNumber == rmaCmDatabaseSheet.getRange(rangeMiddel-1,1).getValue()) //going to the first requested rma row
          { 
            rangeMiddel = rangeMiddel - 1;
          }
          var databaseRmaFirstRow = rangeMiddel;

          var copyRow = 14;
          while (thisSheet.getRange(copyRow,5).getValue() != "") //we didn't go through all the rows in the form
          { 
            var rowExistInDatabase = false;
            rangeMiddel = databaseRmaFirstRow; //starting from the begginig of the rma block
            while (rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue() == rmaNumber) //we didn't go through all the rows with this rma in the database
            {
              if (rmaCmDatabaseSheet.getRange(rangeMiddel,13).getValue() == thisSheet.getRange(copyRow,5).getValue()) //if the luminer in tha data base as the same Comunication card serial number as the one in the technican form
              {
                rowExistInDatabase = true;
                backgroundCode = rmaCmDatabaseSheet.getRange(rangeMiddel,1).getBackground();

                for (var column = 17;column <= 48;column++) //copying all the row data from the databse
                {
                  rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(thisSheet.getRange(copyRow,column - 8).getValue());
                }


                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(rmaDecision);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa1);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa2);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa3);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa4);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa5);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(fixCountry);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(technicianName);
                column++;
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(technicianCode);
                column++;

                var date = Utilities.formatDate(new Date(),"GMT+3:00", 'M/d/yyyy');
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(date);
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('center');

                fixturesQty++; //adding this fixture to the sum of all fixtures in this RMA.
                var currentSerialNumber = thisSheet.getRange(copyRow,6).getValue();
                fixturesSerialNumbers = fixturesSerialNumbers + currentSerialNumber + ","; //adding the current serial number to the RMA serial numbers

                //check if we need to add the current fixtur replaced parts to the array
                if (thisSheet.getRange(copyRow,26).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,26).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,26).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(copyRow,28).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,28).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,28).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(copyRow,30).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,30).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,30).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(copyRow,32).getValue() != "")
                {
                  
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,32).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,32).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(copyRow,35).getValue() != "")
                {
                  
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,35).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,35).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                  
                }

                break;
              }
              rangeMiddel++;
            }


            if (rowExistInDatabase == false) //if the row from the form does not exist in the data base
            {
              rangeMiddel = rmaCmDatabaseSheet.getLastRow() + 1; //the first new row in the database

              //copying all the row data from the databse
              rmaCmDatabaseSheet.getRange(rangeMiddel,1).setValue(rmaNumber);
              rmaCmDatabaseSheet.getRange(rangeMiddel,1).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,1).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,1).setFontWeight('bold');
              rmaCmDatabaseSheet.getRange(rangeMiddel,2).setValue(company);
              rmaCmDatabaseSheet.getRange(rangeMiddel,2).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,2).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,3).setValue(contact);
              rmaCmDatabaseSheet.getRange(rangeMiddel,3).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,3).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,4).setValue(address);
              rmaCmDatabaseSheet.getRange(rangeMiddel,4).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,4).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,5).setValue(city);
              rmaCmDatabaseSheet.getRange(rangeMiddel,5).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,5).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,6).setValue(state);
              rmaCmDatabaseSheet.getRange(rangeMiddel,6).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,6).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,7).setValue(email);
              rmaCmDatabaseSheet.getRange(rangeMiddel,7).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,7).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,8).setValue(phoneNumber);
              rmaCmDatabaseSheet.getRange(rangeMiddel,8).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,8).setHorizontalAlignment('center');
              rmaCmDatabaseSheet.getRange(rangeMiddel,9).setValue(orderDate);
              rmaCmDatabaseSheet.getRange(rangeMiddel,9).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,9).setHorizontalAlignment('center');
              for (var column = 10;column <= 48;column++)
              {
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(thisSheet.getRange(copyRow,column - 8).getValue());
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
                rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('center');
              }
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(rmaDecision);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('left');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa1);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('left');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa2);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('left');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa3);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('left');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa4);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('left');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(Capa5);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('left');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(fixCountry);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('center');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(technicianName);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('center');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(technicianCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setBackground(backgroundCode);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('center');
              column++;
              var date = Utilities.formatDate(new Date(),"GMT+3:00", 'M/d/yyyy');
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setValue(date);
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).setHorizontalAlignment('center');
              column++;
              rmaCmDatabaseSheet.getRange(rangeMiddel,column).insertCheckboxes();
              column++;
              rmaDatabaseSpreadsheet.getSheetByName("analysis").getRange(1,60,1,9).copyTo(rmaCmDatabaseSheet.getRange(rangeMiddel,column),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); //copying from the analysis sheet the cells with the formulas that calculate the costs

              fixturesQty++; //adding this fixture to the sum of all fixtures in this RMA.
              var currentSerialNumber = thisSheet.getRange(copyRow,6).getValue();
                fixturesSerialNumbers = fixturesSerialNumbers + currentSerialNumber + ","; //adding the current serial number to the RMA serial numbers

                //check if we need to add the current fixtur replaced parts to the array
                if (thisSheet.getRange(copyRow,26).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,26).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,26).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                  
                }

                if (thisSheet.getRange(copyRow,28).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,28).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,28).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                  
                }

                if (thisSheet.getRange(copyRow,30).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,30).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,30).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(copyRow,32).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,32).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,32).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(copyRow,35).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null) //going though all the cells in the array
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(copyRow,35).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(copyRow,35).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

            }
            copyRow++;
          }
          break;
        }

        else if (rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue() > rmaNumber) //if the requested rma is smaller than the current middle.
        {
          rangeEnd = rangeMiddel - 1;
        }
        else //if the requested rma is bigger than the current middle.
        {
          rangeStart = rangeMiddel + 1;
        }
      }


      if (rmaExist == false) //if the rma doesn't already exist in the data base we enter the new data in the bottom
      { 
        var filterExist = rmaCmDatabaseSheet.getFilter();
        if (filterExist != null)
        {
          rmaCmDatabaseSheet.getFilter().remove(); //canceling any column filter before pasteing the new values
        }
        //randomly picking a new background color
        var lastBackgroundColorUsedCode = rmaCmDatabaseSheet.getRange(rmaCmDatabaseSheet.getLastRow(),1).getBackground(); //getting the last background color we used, so the new on will be differnt so we can see the differnce.
        var backgroundCode = lastBackgroundColorUsedCode; 
        while (backgroundCode == lastBackgroundColorUsedCode) //we rundomly picking a background color that is differnt from the last one
        {
          var backgroundNumber = Math.floor(Math.random() * 7); //randomly picking a number between 0 to 6
          switch (backgroundNumber) 
          {
            case 0: backgroundCode = "#d9d9d9"; //the hex code for light gray
              break;
            case 1: backgroundCode = "#e6b8af"; //the hex code for light mango
              break;
            case 2: backgroundCode = "#fce5cd"; //the hex code for light orange
              break;
            case 3: backgroundCode = "#fff2cc"; //the hex code for light yellow
              break;
            case 4: backgroundCode = "#cfe2f3"; //the hex code for light blue
              break;
            case 5: backgroundCode = "#d9d2e9"; //the hex code for light purple
              break;
            case 6: backgroundCode = "#d9ead3"; //the hex code for light green
              break;
          }
        }

        var databasecopyRow = rmaCmDatabaseSheet.getLastRow() + 1; //the first new row in the database
        var formStartingRow = 14; //the row we will start copying from
        var i = 0;

        while (thisSheet.getRange(formStartingRow + i,5).isBlank() == false) //going through all the rows with data
        {
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,1).setValue(rmaNumber);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,1).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,1).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,1).setFontWeight('bold');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,2).setValue(company);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,2).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,2).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,3).setValue(contact);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,3).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,3).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,4).setValue(address);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,4).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,4).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,5).setValue(city);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,5).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,5).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,6).setValue(state);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,6).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,6).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,7).setValue(email);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,7).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,7).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,8).setValue(phoneNumber);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,8).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,8).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,9).setValue(orderDate);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,9).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,9).setHorizontalAlignment('center');

          for (var j = 2;j <= 40;j++) //copying the rows from the CM form to the database.
          {
            rmaCmDatabaseSheet.getRange(databasecopyRow + i,j + 8).setValue(thisSheet.getRange(formStartingRow + i,j).getValue());
            rmaCmDatabaseSheet.getRange(databasecopyRow + i,j + 8).setBackground(backgroundCode);
            rmaCmDatabaseSheet.getRange(databasecopyRow + i,j + 8).setHorizontalAlignment('center');

          }

          rmaCmDatabaseSheet.getRange(databasecopyRow + i,49).setValue(rmaDecision);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,49).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,49).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,50).setValue(Capa1);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,50).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,50).setHorizontalAlignment('left');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,51).setValue(Capa2);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,51).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,51).setHorizontalAlignment('left');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,52).setValue(Capa3);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,52).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,52).setHorizontalAlignment('left');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,53).setValue(Capa4);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,53).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,53).setHorizontalAlignment('left');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,54).setValue(Capa5);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,54).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,54).setHorizontalAlignment('left');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,55).setValue(fixCountry);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,55).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,55).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,56).setValue(technicianName);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,56).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,56).setHorizontalAlignment('center');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,57).setValue(technicianCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,57).setBackground(backgroundCode);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,57).setHorizontalAlignment('center');

          var date = Utilities.formatDate(new Date(),"GMT+3:00", 'M/d/yyyy');
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,58).setValue(date);
          rmaCmDatabaseSheet.getRange(databasecopyRow + i,58).setHorizontalAlignment('center');
        rmaCmDatabaseSheet.getRange(databasecopyRow + i,59).insertCheckboxes();

        rmaDatabaseSpreadsheet.getSheetByName("analysis").getRange(1,60,1,9).copyTo(rmaCmDatabaseSheet.getRange(databasecopyRow + i,60),SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); //copying from the first row the cells with the formulas that calculate the costs

        fixturesQty++; //adding this fixture to the sum of all fixtures in this RMA.
                var currentSerialNumber = thisSheet.getRange(formStartingRow + i,6).getValue();
                fixturesSerialNumbers = fixturesSerialNumbers + currentSerialNumber + ","; //adding the current serial number to the RMA serial numbers

                //check if we need to add the current fixtur replaced parts to the array
                if (thisSheet.getRange(formStartingRow + i,26).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(formStartingRow + i,26).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(formStartingRow + i,26).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(formStartingRow + i,28).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(formStartingRow + i,28).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(formStartingRow + i,28).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(formStartingRow + i,30).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(formStartingRow + i,30).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(formStartingRow + i,30).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(formStartingRow + i,32).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(formStartingRow + i,32).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(formStartingRow + i,32).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(formStartingRow + i,35).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(formStartingRow + i,35).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(formStartingRow + i,35).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

                if (thisSheet.getRange(formStartingRow + i,38).getValue() != "")
                {
                    var partExistInArray = false;
                    i = 0;
                    while (replacedPartsSerialNumbers[i] != null)
                    {
                      if (replacedPartsSerialNumbers[i] == thisSheet.getRange(formStartingRow + i,38).getValue())
                      {partExistInArray = true;}
                      i++;
                    }
                    
                    if (partExistInArray == false)
                    {
                      replacedPartsSerialNumbers[replacedPartsArrayLastIndex] = thisSheet.getRange(formStartingRow + i,38).getValue();
                      replacedPartsArrayLastIndex++;
                    }
                }

          i++;

        }

        rmaCmDatabaseSheet.getRange(databasecopyRow,1,i,57).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM); //putting a thick border around the new block of rows.
      }
      SpreadsheetApp.flush(); //making sure all the changes we made in the "CM RMA" are saved before making the summary
      
  //calculating the total repair RMA cost
      var totalRepairCost = 0;

      //preforming a binary search to search the RMA ID
      rmaCmDatabaseSheet.getRange("A:A").setNumberFormat('@'); //Seting the cells format to text so when sorting the column numbers and text will be ordered in the right order and the binary search will work
      var filter = rmaCmDatabaseSheet.getFilter();
      if (filter == null)
        {rmaCmDatabaseSheet.getRange("A:BP").createFilter();}
      rmaCmDatabaseSheet.getFilter().sort(1, true); //sorting the RMA ID column from small to large
      rangeStart = 2;
      rangeEnd = rmaCmDatabaseSheet.getRange("A:A").getLastRow();
      if (rmaCmDatabaseSheet.getRange(rangeEnd,1).getValue() == "") //getting the last row of data
      {
        rmaCmDatabaseSheet.getRange(rangeEnd,1).activate();
        rmaCmDatabaseSheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
        rangeEnd = rmaCmDatabaseSheet.getActiveCell().getRowIndex();
      }

      while (rangeStart <= rangeEnd)//continue the binary search until we coverd all the range
      {
        var rangeMiddel = Math.floor((rangeStart + rangeEnd)/2); //calculating the middle of the range.
        if (rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue() == rmaNumber) //we found the rma
        {
          while (rmaCmDatabaseSheet.getRange(rangeMiddel -1,1).getValue() == rmaNumber) //going to the first row of this rma
          {
            rangeMiddel--;
          }
          while (rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue() == rmaNumber)
          {
            totalRepairCost = totalRepairCost + rmaCmDatabaseSheet.getRange(rangeMiddel,67).getValue(); //adding the current fixture repair cost to the total repair cost
            rangeMiddel++; //moving to the next fixture
          }
          break;
        }
        
        else if (rmaCmDatabaseSheet.getRange(rangeMiddel,1).getValue() > rmaNumber) //if the requested rma is smaller than the current middle.
        {
          rangeEnd = rangeMiddel - 1;
        }
        else //if the requested rma is bigger than the current middle.
        {
          rangeStart= rangeMiddel + 1;
        }
      }


      var rmaSummarySheet = rmaDatabaseSpreadsheet.getSheetByName("RMA Summary");

      if (rmaExist == true) //if the this RMA already exist in the database we will delete the old one
      {
        //preforming a binary search to search the RMA ID
        rmaSummarySheet.getRange("C:C").setNumberFormat('@'); //Seting the cells format to text so when sorting the column numbers and text will be ordered in the right order and the binary search will work
        var filter = rmaSummarySheet.getFilter();
      if (filter == null)
        {rmaSummarySheet.getRange("A:Q").createFilter();}
        rmaSummarySheet.getFilter().sort(3, true); //sorting the RMA ID column from small to large
        rangeStart = 2;
        rangeEnd = rmaSummarySheet.getRange("C:C").getLastRow();
        if (rmaSummarySheet.getRange(rangeEnd,3).getValue() == "") //getting the last row of data
        {
          rmaSummarySheet.getRange(rangeEnd,3).activate();
          rmaSummarySheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
          rangeEnd = rmaSummarySheet.getActiveCell().getRowIndex();
        }

        while (rangeStart <= rangeEnd)//continue the binary search until we coverd all the range
        {
          var rangeMiddel = Math.floor((rangeStart + rangeEnd)/2); //calculating the middle of the range.

          if (rmaNumber == rmaSummarySheet.getRange(rangeMiddel,3).getValue()) //if this the old row
          {
            rmaSummarySheet.deleteRow(rangeMiddel); //deleting the old row with the same RMA ID
            break;
          }

          else if (rmaSummarySheet.getRange(rangeMiddel,1).getValue() > rmaNumber) //if the requested rma is smaller than the current middle.
          {
            rangeEnd = rangeMiddel - 1;
          }
          else //if the requested rma is bigger than the current middle.
          {
            rangeStart= rangeMiddel + 1;
          }
        }
      }

      //creating the RMA summary raw
      rmaSummarySheet.getRange("B:B").setNumberFormat('@'); //Seting the cells format to text so when sorting the column numbers and text will be ordered in the right order and the binary search will work
      var filter = rmaSummarySheet.getFilter();
      if (filter == null)
        {rmaSummarySheet.getRange("A:Q").createFilter();}
  rmaSummarySheet.getFilter().sort(2, true); //sorting the RMA # column from small to large
      var summarySheetLastRow = rmaSummarySheet.getLastRow(); //getting the last row
      if (rmaSummarySheet.getRange(summarySheetLastRow,3).getValue() == "") //getting the last row of data
        {
          rmaSummarySheet.getRange(summarySheetLastRow,3).activate();
          rmaSummarySheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
          summarySheetLastRow = rmaSummarySheet.getActiveCell().getRowIndex();
        }
      var summaryNewRow = summarySheetLastRow + 1;

      rmaSummarySheet.getRange(summaryNewRow,1).insertCheckboxes();
      
      var rmaOrderNumberCount = Number(rmaSummarySheet.getRange(summarySheetLastRow,2).getValue());
      if (rmaOrderNumberCount == "#") {rmaOrderNumberCount = 0}
      rmaOrderNumberCount = rmaOrderNumberCount + 1;
      rmaSummarySheet.getRange(summaryNewRow,2).setValue(rmaOrderNumberCount);
      rmaSummarySheet.getRange(summaryNewRow,2).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,2).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,3).setValue(rmaNumber);
      rmaSummarySheet.getRange(summaryNewRow,3).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,3).setBackground(backgroundCode);

      var pn = thisSheet.getRange(14,2).getValue();
      rmaSummarySheet.getRange(summaryNewRow,4).setValue(pn);
      rmaSummarySheet.getRange(summaryNewRow,4).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,4).setBackground(backgroundCode);
      var fixture = thisSheet.getRange(14,3).getValue();
      rmaSummarySheet.getRange(summaryNewRow,5).setValue(fixture);
      rmaSummarySheet.getRange(summaryNewRow,5).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,5).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,6).setValue(fixturesQty);
      rmaSummarySheet.getRange(summaryNewRow,6).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,6).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,7).setValue(fixturesSerialNumbers);
      rmaSummarySheet.getRange(summaryNewRow,7).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,7).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,8).setValue("");
      rmaSummarySheet.getRange(summaryNewRow,8).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,8).setBackground(backgroundCode);
      var arrivalDate = thisSheet.getRange(14,8).getValue();
      rmaSummarySheet.getRange(summaryNewRow,9).setValue(arrivalDate);
      rmaSummarySheet.getRange(summaryNewRow,9).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,9).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,10).setValue(date);
      rmaSummarySheet.getRange(summaryNewRow,10).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,10).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,11).setValue(company);
      rmaSummarySheet.getRange(summaryNewRow,11).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,11).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,12).setValue(city);
      rmaSummarySheet.getRange(summaryNewRow,12).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,12).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,13).setValue(contact);
      rmaSummarySheet.getRange(summaryNewRow,13).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,13).setBackground(backgroundCode);
      var failureType = thisSheet.getRange(14,22).getValue();
      rmaSummarySheet.getRange(summaryNewRow,14).setValue(failureType);
      rmaSummarySheet.getRange(summaryNewRow,14).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,14).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,15).setValue(replacedPartsSerialNumbers);
      rmaSummarySheet.getRange(summaryNewRow,15).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,15).setBackground(backgroundCode);
      rmaSummarySheet.getRange(summaryNewRow,16).setValue(totalRepairCost);
      rmaSummarySheet.getRange(summaryNewRow,16).setHorizontalAlignment('center');
      rmaSummarySheet.getRange(summaryNewRow,16).setBackground(backgroundCode);

      //creating the RMA summary file and save it in folder
      var newSpreadsheetName = company + " - " + rmaNumber; //naming the new file
      var folderId = '1T8VrTriNMNCmX0SuhyRXjcoyPfZrJnmq'
      var resource = {
      title: newSpreadsheetName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: folderId }]
      }
      var temporarySpreadsheet = Drive.Files.insert(resource) //creating the new spreadsheet in the folder
      var temporarySpreadsheetId = temporarySpreadsheet.id; //getting the new spreadsheet id.

      //pasting the rows to the new spreassheet
      var newSpreadsheet = SpreadsheetApp.openById(temporarySpreadsheetId);
      var newSpreadsheetSheet = newSpreadsheet.getSheets()[0];
      var copyRange = rmaSummarySheet.getRange(1,2,1,16).getValues();
      newSpreadsheetSheet.getRange(1,1,1,16).setValues(copyRange);
      copyRange = rmaSummarySheet.getRange(1,2,1,16).getBackgrounds();
      newSpreadsheetSheet.getRange(1,1,1,16).setBackgrounds(copyRange);
      newSpreadsheetSheet.getRange(1,1,1,16).setFontWeight('bold');
      copyRange = rmaSummarySheet.getRange(summaryNewRow,2,1,16).getValues();
      newSpreadsheetSheet.getRange(2,1,1,16).setValues(copyRange);
      newSpreadsheetSheet.getRange(1,1,2,16).setHorizontalAlignment('center');

      SpreadsheetApp.flush(); //making sure the changes we made in the new spreadsheet are beeing saved







  SpreadsheetApp.getUi().alert("The data transformed to the database successfully!");

  //sending an alert email to the mananger
  var url = 'https://docs.google.com/spreadsheets/d/'+ temporarySpreadsheetId +'/export?format=xlsx';
  var token         = ScriptApp.getOAuthToken();
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });

  var fileName = newSpreadsheetName + '.xlsx';
  var blobs   = [response.getBlob().setName(fileName)];
  var managerEmail = rmaDatabaseSpreadsheet.getSheetByName("CUSTOMERS RMA").getRange("copyDefultEmail").getValue();
  MailApp.sendEmail(managerEmail + ",adik@juganu.com","Technican RMA update","Juganu's technican (" +technicianName+ ") as submited a new RMA into the database (RMA:" +rmaNumber+ ").",{attachments: blobs});



}











//A function that cleans all the cells in the form
function cleanAllCells()
{
  var buttonPressed = SpreadsheetApp.getUi().alert("Are you sure you want to clean all the cells in the form? it can take a minute.",SpreadsheetApp.getUi().ButtonSet.YES_NO);

  if (buttonPressed == SpreadsheetApp.getUi().Button.NO || buttonPressed == SpreadsheetApp.getUi().Button.CLOSE) //if the user decided not to continue
  {
    return;
  }

  var thisSheet = SpreadsheetApp.getActiveSheet();
//cleaning all the cells
  thisSheet.getRange("RMA").setValue("");
  thisSheet.getRange("Company").setValue("");
  thisSheet.getRange("Contact").setValue("");
  thisSheet.getRange("Address").setValue("");
  thisSheet.getRange("City").setValue("");
  thisSheet.getRange("State").setValue("");
  thisSheet.getRange("Email").setValue("");
  thisSheet.getRange("phoneNumber").setValue("");
  thisSheet.getRange("OrderDate").setValue("");
  thisSheet.getRange("rmaDecision").setValue("");
  thisSheet.getRange("fixCountry").setValue("");
  thisSheet.getRange("capa").setValue("");
  thisSheet.getRange("technicianName").setValue("");
  thisSheet.getRange("B14:J1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  for (var i = 14;i <= 1999;i++)
  {
    for (var j = 11;j <= 19;j++)
    {
      thisSheet.getRange(i,j).setValue("FALSE");
    }
  }

  thisSheet.getRange("T14:V1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  thisSheet.getRange("X14:X1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  for (var i = 14;i <= 1999;i++)
    {
      thisSheet.getRange(i,25).setValue("FALSE");
    }

  thisSheet.getRange("Z14:Z1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});  

  for (var i = 14;i <= 1999;i++)
    {
      thisSheet.getRange(i,27).setValue("FALSE");
    }  

  thisSheet.getRange("AB14:AB1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});    

  for (var i = 14;i <= 1999;i++)
    {
      thisSheet.getRange(i,29).setValue("FALSE");
    }  

  thisSheet.getRange("AD14:AD1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});   

  for (var i = 14;i <= 1999;i++)
    {
      thisSheet.getRange(i,31).setValue("FALSE");
    }   

  thisSheet.getRange("AF14:AG1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true}); 

  for (var i = 14;i <= 1999;i++)
    {
      thisSheet.getRange(i,34).setValue("FALSE");
    }   

  thisSheet.getRange("AI14:AJ1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true}); 

  for (var i = 14;i <= 1999;i++)
    {
      thisSheet.getRange(i,37).setValue("FALSE");
    }   

  thisSheet.getRange("AL14:AM1999").activate();
  thisSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  for (var i = 14;i <= 1999;i++)
    {
      thisSheet.getRange(i,40).setValue("FALSE");
    }   
}
