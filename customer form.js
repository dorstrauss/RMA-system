/** a function that execute when the user press the "submit" button.
 * the function copys all the data that the user insert in the form, to another spreedsheet the work as a data-base.
 */
function customerRmaForm() {
  var customerFormSheet = SpreadsheetApp.getActiveSheet();

  //geting the values from the form
  var company = customerFormSheet.getRange("company").getValue();
  var contact = customerFormSheet.getRange("contact").getValue();
  var address = customerFormSheet.getRange("address").getValue();
  var phoneNumber = customerFormSheet.getRange("phoneNumber").getValue();
  var city = customerFormSheet.getRange("city").getValue();
  var state = customerFormSheet.getRange("state").getValue();
  var zip = customerFormSheet.getRange("zip").getValue();
  var email = customerFormSheet.getRange("email").getValue();
  var totalQty = customerFormSheet.getRange("totalQty").getValue();
  var orderDate = customerFormSheet.getRange("orderDate").getValue();

  //checking if all the fields are filled correctly
  var formFilled = true;
  if (company == "" || contact == "" || address == "" || phoneNumber == "" || city == "" || state == "" || zip == "" || email == "")
  {
    formFilled = false;
  }
  if (formFilled == false)
  {
    SpreadsheetApp.getUi().alert("One or more than the fields in the form is empty, please make sure to fill all the fields before submiting.");
    return;
  }

  

  //getting the current form subscription number, and raising it by 1.
  var currentFormNumber = customerFormSheet.getRange("subscriptionNumber").getValue();
  customerFormSheet.getRange("subscriptionNumber").setValue(currentFormNumber + 1);

  //opening the database sheet where the customers form values are stored
  var rmaDatabaseSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1B3PIk0T0F4txYTE9daP1pVCdhH2t_zVGd6Uz_SLuZi4/edit?usp=sharing");
  var rmaDatabaseSheet = rmaDatabaseSpreadsheet.getSheetByName("CUSTOMERS RMA");

  //copying all the values from the form into the database
  var filter = rmaDatabaseSheet.getFilter();
  if (filter == null)
    {rmaDatabaseSheet.getRange("A:U").createFilter();}
  rmaDatabaseSheet.getFilter().sort(2, true); //sorting the form submission numbers in ascending order
  var formFirstRow = 15; //the first row in the form table
  var formFirstColumn = 3; //the first column in the form table
  var databaseFirstRow = rmaDatabaseSheet.getLastRow() + 1; //the first new row in the database
  var databseFirstColumn = 3;
  var formRowsCounter = 0;
  var databaseRowCounter = 0;
  
  //checking if the user trying to submite products that already exist in the database.
  var stopCode = false;
  while (customerFormSheet.getRange(formFirstRow + formRowsCounter, 3).isBlank() == false)
  {
    while (rmaDatabaseSheet.getRange(2 + databaseRowCounter,3).isBlank() == false)
    {
      if (customerFormSheet.getRange(formFirstRow + formRowsCounter, 6).getValue() == rmaDatabaseSheet.getRange(2 + databaseRowCounter,6).getValue()) //if this Comunication serial number already exist in the data base
      {
        var buttonPressed = SpreadsheetApp.getUi().alert("In the form there is one or more Communication serial number that already exists in the database, are you sure you want to continue? ",SpreadsheetApp.getUi().ButtonSet.YES_NO);  //desplaying a message if there are a P.N that already exist in the data base.

          if (buttonPressed == SpreadsheetApp.getUi().Button.NO) //if the user preesed NO it will stop the code.
          {
            return;
          }
          else if (buttonPressed == SpreadsheetApp.getUi().Button.YES) //if the user choose YES, it will end the while loop
          {
            stopCode = true; //we change the value to true so when we will reach the end of the outer while loop, the outer while loop will stop
            break;
          } 
          else {return;} //if the user press the "X" button it will stop the cose.
      }
      databaseRowCounter++;
    }
    if (stopCode == true) //if the user chose to contiue the code the while loop ends
      {break;}
    formRowsCounter++;
  }

  //randomly picking a new background color
  var lastBackgroundColorUsedCode = rmaDatabaseSheet.getRange(rmaDatabaseSheet.getLastRow(),2).getBackground(); //getting the last background color we used, so the new on will be differnt so we can see the differnce.
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

  formRowsCounter = 0;
  while (customerFormSheet.getRange(formFirstRow + formRowsCounter, 3).isBlank() == false) //keeps running until a empty P.N row
  {
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, 2).setValue(currentFormNumber); //writing the form number
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, 2).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, 2).setHorizontalAlignment('center');

    for (var columnsCounter = 0; columnsCounter < 6; columnsCounter++) {
      var copyValue = customerFormSheet.getRange(formFirstRow + formRowsCounter, formFirstColumn + columnsCounter).getValue(); //copying the value from the form
      rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(copyValue); //pasting the value to the database
      rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
      rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    }

    //puting the other values from the form in each row
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(company);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(contact);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(address);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(phoneNumber);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(city);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(state);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(zip);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(email);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(totalQty);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setValue(orderDate);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setBackground(backgroundCode);
    rmaDatabaseSheet.getRange(databaseFirstRow + formRowsCounter, databseFirstColumn + columnsCounter).setHorizontalAlignment('center');
    columnsCounter++;

    formRowsCounter = formRowsCounter + 1; //moving to the next row
  }

  rmaDatabaseSheet.getRange(databaseFirstRow,2,formRowsCounter,17).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM); //putting a thick border around the new block of rows.

  //Deleting the customer data after he clicks submit
  // customerFormSheet.getRange("company").setValue("");
  // customerFormSheet.getRange("contact").setValue("");
  // customerFormSheet.getRange("address").setValue("");
  // customerFormSheet.getRange("phoneNumber").setValue("");
  // customerFormSheet.getRange("state").setValue("");
  // customerFormSheet.getRange("city").setValue("");
  // customerFormSheet.getRange("zip").setValue("");
  // customerFormSheet.getRange("email").setValue("");
  // customerFormSheet.getRange("totalQty").setValue("");
  // customerFormSheet.getRange("orderDate").setValue("");
  
  //getting the last row of data
  // customerFormSheet.getRange(1000,3).activate();
  // customerFormSheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  // var dataLastRow = customerFormSheet.getActiveRange().getRow();

  // customerFormSheet.getRange(15,3,985,1).clear({contentsOnly: true, skipFilteredRows: true});
  // customerFormSheet.getRange(15,5,985,4).clear({contentsOnly: true, skipFilteredRows: true});

  // customerFormSheet.getRange("company").activate();

  //creating a new spreadsheet with the new RMA data in order to sent by email to the customer
  var newSpreadsheetName = company + " RMA Request";
  var folderId = '1wb5rA4DrmdmEpW-03X04G3wKs3SkSpEa'
  var resource = {
  title: newSpreadsheetName,
  mimeType: MimeType.GOOGLE_SHEETS,
  parents: [{ id: folderId }]
  }
  var temporarySpreadsheet = Drive.Files.insert(resource) //creating the new spreadsheet in the folder
  var temporarySpreadsheetId = temporarySpreadsheet.id; //getting the new spreadsheet id.
  var newSpreadsheet = SpreadsheetApp.openById(temporarySpreadsheetId);
  var newSpreadsheetSheet = newSpreadsheet.getSheets()[0];
  
  //creating the headers
  newSpreadsheetSheet.getRange(1,7).setValue("Company");
  newSpreadsheetSheet.getRange(1,7).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,7).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,8).setValue("Contact");
  newSpreadsheetSheet.getRange(1,8).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,8).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,9).setValue("Address");
  newSpreadsheetSheet.getRange(1,9).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,9).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,10).setValue("Phone");
  newSpreadsheetSheet.getRange(1,10).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,10).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,11).setValue("City");
  newSpreadsheetSheet.getRange(1,11).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,11).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,12).setValue("State");
  newSpreadsheetSheet.getRange(1,12).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,12).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,13).setValue("Zip");
  newSpreadsheetSheet.getRange(1,13).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,13).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,14).setValue("Email");
  newSpreadsheetSheet.getRange(1,14).setBackground('#6fa8dc');
  newSpreadsheetSheet.getRange(1,14).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,1).setValue("P.N");
  newSpreadsheetSheet.getRange(1,1).setBackground('green');
  newSpreadsheetSheet.getRange(1,1).setFontWeight('bold');
   newSpreadsheetSheet.getRange(1,2).setValue("Item");
  newSpreadsheetSheet.getRange(1,2).setBackground('green');
  newSpreadsheetSheet.getRange(1,2).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,3).setValue("Reason for Return");
  newSpreadsheetSheet.getRange(1,3).setBackground('green');
  newSpreadsheetSheet.getRange(1,3).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,4).setValue("Quantity");
  newSpreadsheetSheet.getRange(1,4).setBackground('green');
  newSpreadsheetSheet.getRange(1,4).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,5).setValue("S.N Luminaire");
  newSpreadsheetSheet.getRange(1,5).setBackground('green');
  newSpreadsheetSheet.getRange(1,5).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,6).setValue("Order Date");
  newSpreadsheetSheet.getRange(1,6).setBackground('green');
  newSpreadsheetSheet.getRange(1,6).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,1,1,15).setHorizontalAlignment("center");

  var currentCopyRow = 15;
  var currentPasteRow = 2;
  while (customerFormSheet.getRange(currentCopyRow,3).getValue() != "")
  {
    newSpreadsheetSheet.getRange(currentPasteRow,5).setNumberFormat("@"); //set the cell to text
    newSpreadsheetSheet.getRange(currentPasteRow,1,1,6).setValues(customerFormSheet.getRange(currentCopyRow,3,1,6).getValues());
    newSpreadsheetSheet.getRange(currentPasteRow,7).setValue(company);
    newSpreadsheetSheet.getRange(currentPasteRow,8).setValue(contact);
    newSpreadsheetSheet.getRange(currentPasteRow,9).setValue(address);
    newSpreadsheetSheet.getRange(currentPasteRow,10).setValue(phoneNumber);
    newSpreadsheetSheet.getRange(currentPasteRow,11).setValue(city);
    newSpreadsheetSheet.getRange(currentPasteRow,12).setValue(state);
    newSpreadsheetSheet.getRange(currentPasteRow,13).setValue(zip);
    newSpreadsheetSheet.getRange(currentPasteRow,14).setValue(email);
    newSpreadsheetSheet.getRange(currentPasteRow,1,1,14).setHorizontalAlignment("center");

    currentCopyRow++;
    currentPasteRow++;
  }

  SpreadsheetApp.flush(); //making sure the changes we made in the new spreadsheet are beeing saved

  //sending an email with the new RMA data to the manager
  var url = 'https://docs.google.com/spreadsheets/d/'+ temporarySpreadsheetId +'/export?format=xlsx';
  var token         = ScriptApp.getOAuthToken();
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  var fileName = newSpreadsheetName + '.xlsx';
  var blobs   = [response.getBlob().setName(fileName)];

  MailApp.sendEmail("juganurmasystem@gmail.com,adik@juganu.com","New RMA request from " + company,"A customer submitted a new RMA request, the full details are in the RMA database.", {attachments: blobs});

  SpreadsheetApp.getUi().alert('Your RMA request has been received.' + "\n" + 'Once Juganu will approve your request you will get an email with your RMA trucking number to the Email address you entered in the form.' + "\n" + "Please reply to the email and attach an images of the fixtures packed before delivery.");

}








/** a function that execute when someone opens the sheet/form */
function onOpen()
{
  var customerFormSheet = SpreadsheetApp.getActiveSheet();
  customerFormSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  //Deleting the values from the previus customer that filled the form
  customerFormSheet.getRange("company").setValue("");
  customerFormSheet.getRange("contact").setValue("");
  customerFormSheet.getRange("address").setValue("");
  customerFormSheet.getRange("phoneNumber").setValue("");
  customerFormSheet.getRange("state").setValue("");
  customerFormSheet.getRange("city").setValue("");
  customerFormSheet.getRange("zip").setValue("");
  customerFormSheet.getRange("email").setValue("");
  customerFormSheet.getRange("totalQty").setValue("");
  customerFormSheet.getRange("orderDate").setValue("");
  customerFormSheet.getRange(15,3,985,1).clear({contentsOnly: true, skipFilteredRows: true});
  customerFormSheet.getRange(15,5,985,4).clear({contentsOnly: true, skipFilteredRows: true});

  customerFormSheet.getRange(6,7).setValue("Phone");

  customerFormSheet.getRange("company").activate();

  SpreadsheetApp.getUi().alert("Thank you for contacting Juganu, please fill in the RMA form and click submit");

}
