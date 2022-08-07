/** a function that is being called when the operator of the system wants to give a uniqe RMA number to a certen rma
 * request that was filled in the form by the customer.
 */
function rmaGenerator()
{ 
  var databaseSheet = SpreadsheetApp.getActiveSheet();
  var activeRangesList = databaseSheet.getActiveRangeList();
  var activeRanges = activeRangesList.getRanges(); // getting a list (array of arrays) of the selected ranges
  
  //going through the selected rows to check if there are different form submisions number, if so a massage will apper and ask the user if he is sure he want to continue
  var previusFormId = -9999; //setting a value before the first loop
  for (var i = 0;i < activeRanges.length;i++)
  {
    var currentRange = activeRanges[i]; //set the range we going through between the selected ranges
    var currentRangeStartingRow = currentRange.getRow();

    for (var r = 0;r < currentRange.getNumRows();r++)
    {
      if (databaseSheet.getRange(currentRangeStartingRow + r,2).getValue() != previusFormId && previusFormId != -9999) //if two neighbors have different form submition id
      {
        var buttonPressed = SpreadsheetApp.getUi().alert("The rows you selected contain data from different form submissions, are you sure you want to give them an RMA number? ",SpreadsheetApp.getUi().ButtonSet.YES_NO);  //desplaying a message if ther are in the rabge selected 2 rows that have different form id number, and saving the reply of the user

        if (buttonPressed == SpreadsheetApp.getUi().Button.NO) //if the user preesed NO it will stop the code.
        {
          return;
        }
        else if (buttonPressed == SpreadsheetApp.getUi().Button.YES) //if the user choose YES, it will end the for loop
        {
          i = activeRanges.length;
          break;
        } 
        else {return;}
      }
      previusFormId = databaseSheet.getRange(currentRangeStartingRow + r,2).getValue();
    }
    
  }

  for (var i = 0;i < activeRanges.length;i++) //going through the selected row to check if one or more already have a RMA ID.
  {
    var currentRange = activeRanges[i]; //set the range we going through between the selected ranges
    var currentRangeStartingRow = currentRange.getRow();

    for (var r = 0;r < currentRange.getNumRows();r++)
    {
      if (databaseSheet.getRange(currentRangeStartingRow + r,1).isBlank() == false) //if there is already RMA
      {
        var buttonPressed = SpreadsheetApp.getUi().alert("One or more of the rows you selected already has an RMA ID, are you sure you want to give them a new RMA ID?",SpreadsheetApp.getUi().ButtonSet.YES_NO);

        if (buttonPressed == SpreadsheetApp.getUi().Button.NO) //if the user don't want to continue
        {
          return;
        }
        else if (buttonPressed == SpreadsheetApp.getUi().Button.YES) //if the user want to continue there is no point in keep checking RMA ID in other rows
        {
          i = activeRanges.length;
          break;
        }
        else {return;}
      }
    }
  }

  //getting the customer email adress and check if it valid
  var inputCustomerEmail = Browser.inputBox("The customer default Email to be sent is: " + databaseSheet.getRange(activeRanges[0].getRow(),16).getValue() + "\\n" + "If you want the Email to be sent to another address, put the address in the input box. otherwise, leave it blank.");
  if (inputCustomerEmail == "cancel") {return;} //if the user click the X of the message box
  else if (inputCustomerEmail == "")
    var customerEmail = databaseSheet.getRange(activeRanges[0].getRow(),16).getValue();
  else
    var customerEmail = inputCustomerEmail;
  

  var inputCopyEmail = Browser.inputBox("The copy of this RMA will be sent to: " + databaseSheet.getRange("copyDefultEmail").getValue() + "\\n" + "If you want the Email to be sent to another address, put the address in the input box. otherwise, leave it blank." );
  if (inputCopyEmail == "cancel") {return;} //if the user click the X of the message box
  else if (inputCopyEmail == "")
    var copyEmail = databaseSheet.getRange("copyDefultEmail").getValue();
  else
    var copyEmail = inputCopyEmail;
  emailCheck = validateEmail(copyEmail); //checking if the email the user entered is valid
  if (emailCheck == false) //if the email address the user entered is invalid
  {
    Browser.msgBox("The copy Email address you entered is invalid!");
    return;
  }

  //get the uniqe RMA id from the user and set it in the selected rows
  var rmaNumber = Browser.inputBox("Enter the new uniqe RMA number you want to give the rows you selected");
  if (rmaNumber == "cancel") {return;} //if the user clicked "X" the code will stop

  //checking if this RMA number already exists in the data base
  //databaseSheet.getFilter().sort(1, true); //canceling columns filters before going through the rows
  i = 2;
  var databaseLastRow = databaseSheet.getLastRow();
  while (i <= databaseLastRow)
  {
    if (databaseSheet.getRange(i,1).getValue() == rmaNumber) //if the rma already exist
    {
        var buttonPressed = SpreadsheetApp.getUi().alert("The new RMA number you entered already exist in the database, are you sure you want to continue?",SpreadsheetApp.getUi().ButtonSet.YES_NO);

        if (buttonPressed == SpreadsheetApp.getUi().Button.NO) //if the user don't want to continue
        {
          return;
        }
        else if (buttonPressed == SpreadsheetApp.getUi().Button.YES) //if the user want to continue there is no point in keep checking RMA ID in other rows
        {
          i = activeRanges.length;
          break;
        }
        else {return;}
    }
    i++;
  }

  
  var newSpreadsheetName = databaseSheet.getRange(activeRanges[0].getRow(),9).getValue() + " - " + rmaNumber;

  for (var i = 0;i < activeRanges.length;i++) //setting the new RMA number
  {
    var currentRange = activeRanges[i]; //set the range we going through between the selected ranges
    var currentRangeStartingRow = currentRange.getRow();

    for (var r = 0;r < currentRange.getNumRows();r++)
    {
      databaseSheet.getRange(currentRangeStartingRow + r,1).setValue(rmaNumber);
      databaseSheet.getRange(currentRangeStartingRow + r,1).setFontWeight('bold');
      databaseSheet.getRange(currentRangeStartingRow + r,1).setHorizontalAlignment('center');

    }
  }


  //creating a new spreadsheet with the new RMA data in order to sent by email to the customer
  var folderId = '1wb5rA4DrmdmEpW-03X04G3wKs3SkSpEa'
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

  newSpreadsheetSheet.getRange(1,1).setValue(databaseSheet.getRange(1,1).getValue()); //pasting the first row header
  newSpreadsheetSheet.getRange(1,1).setBackground(databaseSheet.getRange(1,1).getBackground());
  newSpreadsheetSheet.getRange(1,1).setHorizontalAlignment(databaseSheet.getRange(1,1).getHorizontalAlignment());

  var currentPasteRow = 1;
  var currentCopyRow = databaseSheet.getRange(1,3,1,16).getValues(); //copying the first row values
  newSpreadsheetSheet.getRange(currentPasteRow,2,1,16).setValues(currentCopyRow);
  currentCopyRow = databaseSheet.getRange(1,3,1,16).getBackgrounds(); //copying the first row backgrounds
  newSpreadsheetSheet.getRange(currentPasteRow,2,1,16).setBackgrounds(currentCopyRow);
  currentCopyRow = databaseSheet.getRange(1,3,1,16).getHorizontalAlignments(); //copying the first row horizontal alignments
  newSpreadsheetSheet.getRange(currentPasteRow,2,1,16).setHorizontalAlignments(currentCopyRow);
  newSpreadsheetSheet.getRange(1,1,1,17).setFontWeight('bold');
  newSpreadsheetSheet.getRange(1,1,1,17).setFontFamily('Inter');
  currentPasteRow++;

  //copying and pasting the other rows to the new spread sheet
  for (var i = 0;i < activeRanges.length;i++)
  {
    var currentRange = activeRanges[i]; //set the range we going through between the selected ranges
    var currentRangeStartingRow = currentRange.getRow();
    var company = databaseSheet.getRange(currentRangeStartingRow,9).getValue();

    for (var r = 0;r < currentRange.getNumRows();r++)
    {
      newSpreadsheetSheet.getRange(currentPasteRow,1).setValue(databaseSheet.getRange(currentRangeStartingRow + r,1).getValue()); //pasting the current row RMA number
      newSpreadsheetSheet.getRange(currentPasteRow,1).setHorizontalAlignment('center');
      newSpreadsheetSheet.getRange(currentPasteRow,1).setFontWeight('bold'); //making the RMA ID bold
      newSpreadsheetSheet.getRange(currentPasteRow,1).setFontFamily('Inter');

      currentCopyRow = databaseSheet.getRange(currentRangeStartingRow + r,3,1,16).getValues();
      newSpreadsheetSheet.getRange(currentPasteRow,2,1,16).setValues(currentCopyRow);
      currentCopyRow = databaseSheet.getRange(currentRangeStartingRow + r,3,1,16).getBackgrounds();
      newSpreadsheetSheet.getRange(currentPasteRow,2,1,16).setBackgrounds(currentCopyRow);
      currentCopyRow = databaseSheet.getRange(currentRangeStartingRow + r,3,1,16).getHorizontalAlignments();
      newSpreadsheetSheet.getRange(currentPasteRow,2,1,16).setHorizontalAlignments(currentCopyRow);
      newSpreadsheetSheet.getRange(currentPasteRow,1,1,17).setFontFamily('Inter');

      newSpreadsheetSheet.getRange(currentPasteRow,16).setValue(1); //setting the quantity of each line to be 1
      currentPasteRow++;
    }
  }

  newSpreadsheetSheet.getRange(1,1,currentPasteRow - 1,17).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  newSpreadsheetSheet.getRange(1,1,currentPasteRow - 1,17).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  SpreadsheetApp.flush(); //making sure the changes we made in the new spreadsheet are beeing saved

  //sending an email with the new RMA number and RMA file attached to the customer and a copy to the manager
  var url = 'https://docs.google.com/spreadsheets/d/'+ temporarySpreadsheetId +'/export?format=xlsx';
  var token         = ScriptApp.getOAuthToken();
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  var fileName = newSpreadsheetName + '.xlsx';
  var blobs   = [response.getBlob().setName(fileName)];
  MailApp.sendEmail(customerEmail + "," + copyEmail, company + ' RMA tracking number: ' +rmaNumber,  "Dear " + company + "," + "\n" + 'Your request is being processed by JUGANU, RMA tracking number: ' + rmaNumber + "\n" + "Please go over the attached file to see all the details you entered in the form are correct." + "\n" + "Before shipping the RMA, please send all documents (packing slip and RMA invoice) to Juganu logistic department through email: Hanak@juganu.com" + "\n" + "Please reply to this email and attach the images of the fixtures before delivery." + "\n" + "\n" + "Sincerely," + "\n" + "Adi Kaplan, Head of Quality & Security", {attachments: blobs});

  Browser.msgBox("The new RMA process Succeeded, an email has been sent to the customer and a copy email has been sent to you.");
}












//a function that execute when the technician finished to fix the fixtures and retrived the data to the database. then the manager highlight the specific rows of a rma that he want to ship back to the customer, an email is sent to the logistics 
function managerFixApproval()
{
  var databaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CM RMA");
  var activeRanges = databaseSheet.getActiveRangeList().getRanges(); // getting a list (array of arrays) of the selected ranges
  
  //going through the selected rows to check if there are different RMA ids, if so a massage will apper and ask the user if he is sure he want to continue
  var previusRmaId = -9999; //setting a value before the first loop
  for (var i = 0;i < activeRanges.length;i++) //going through the different ranges
  {
    var currentRange = activeRanges[i]; //set the range we going through between the selected ranges
    var currentRangeStartingRow = currentRange.getRow();

    for (var r = 0;r < currentRange.getNumRows();r++) //going through the rows in the current range
    {
      if (databaseSheet.getRange(currentRangeStartingRow + r,1).getValue() != previusRmaId && previusRmaId != -9999) //if two neighbors have different RMA id
      {
        var buttonPressed = SpreadsheetApp.getUi().alert("The rows you selected contain fixtures with different RMA ids, are you sure you want to continue? ",SpreadsheetApp.getUi().ButtonSet.YES_NO);  //desplaying a message if ther are in the range selected 2 rows that have different rma id, and saving the reply of the user

        if (buttonPressed == SpreadsheetApp.getUi().Button.NO) //if the user preesed NO it will stop the code.
        {
          return;
        }
        else if (buttonPressed == SpreadsheetApp.getUi().Button.YES) //if the user choose YES, it will end the for loop
        {
          i = activeRanges.length;
          break;
        } 
        else {return;}
      }
      previusRmaId = databaseSheet.getRange(currentRangeStartingRow + r,1).getValue();
    }
    
  }

  //getting the logistics email adress and check if it valid
  var copyEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CUSTOMERS RMA').getRange("copyDefultEmail").getValue(); 
  var defaultLogisticsEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CUSTOMERS RMA').getRange("logisticsEmail").getValue(); 
  var inputLogisticsEmail = Browser.inputBox("The logistics default Email to be sent is: " + defaultLogisticsEmail + "\\n" + "If you want the Email to be sent to another address, put the address in the input box. otherwise, leave it blank.");
  if (inputLogisticsEmail == "cancel") {return;} //if the user click the X of the message box
  else if (inputLogisticsEmail != "") //if the user entered a email address.
    defaultLogisticsEmail = inputLogisticsEmail;

  var newSpreadsheetName = Browser.inputBox("Enter the name of the copied spreadsheet that will be emailed to the logistics."); //getting from the user the name he want to give the new spreadsheet.
  if (newSpreadsheetName == "cancel") {return;} //if the user pressed the X button the code will stop.
  //creating a new spreadsheet with the new RMA data in order to sent by email to the logistics
  var folderId = '1UEWksV9xYCuNgaIcwMZYE5oVHtJb1Bx7'
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

  var currentPasteRow = 1;
  var currentCopyRow = databaseSheet.getRange(1,1,1,16).getValues(); //copying the first row values
  newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setValues(currentCopyRow);
  currentCopyRow = databaseSheet.getRange(1,1,1,16).getBackgrounds(); //copying the first row backgrounds
  newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setBackgrounds(currentCopyRow);
  currentCopyRow = databaseSheet.getRange(1,1,1,16).getHorizontalAlignments(); //copying the first row horizontal alignments
  newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setHorizontalAlignments(currentCopyRow);
  for (var i = 1;i <= 17;i++)
  {
    newSpreadsheetSheet.getRange(1,i).setFontWeight('bold');
  }
  newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  currentPasteRow++;

  //copying and pasting the other rows to the new spreadsheet
  for (var i = 0;i < activeRanges.length;i++) //going through the different ranges the managare selected
  {
    var currentRange = activeRanges[i]; //set the range we going through between the selected ranges
    var currentRangeStartingRow = currentRange.getRow();

    for (var r = 0;r < currentRange.getNumRows();r++) //going through the rows in the current range
    {
      currentCopyRow = databaseSheet.getRange(currentRangeStartingRow + r,1,1,16).getValues();
      newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setValues(currentCopyRow);
      currentCopyRow = databaseSheet.getRange(currentRangeStartingRow + r,1,1,16).getBackgrounds();
      newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setBackgrounds(currentCopyRow);
      currentCopyRow = databaseSheet.getRange(currentRangeStartingRow + r,1,1,16).getHorizontalAlignments();
      newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setHorizontalAlignments(currentCopyRow);
      newSpreadsheetSheet.getRange(currentPasteRow,1,1,16).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      currentPasteRow++;
    }
  }
  
  SpreadsheetApp.flush(); //making sure the changes we made in the new spreadsheet are beeing saved  

  //geting the company name
  var firstRange = activeRanges[0];
  var firstRangeFirstRow = firstRange.getRow();
  var company = databaseSheet.getRange(firstRangeFirstRow,2).getValue();

  //sending an email with the new RMA file attached to the logistics and a copy to the manager
  var url = 'https://docs.google.com/spreadsheets/d/'+ temporarySpreadsheetId +'/export?format=xlsx';
  var token         = ScriptApp.getOAuthToken();
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  var fileName = newSpreadsheetName + '.xlsx';
  var blobs   = [response.getBlob().setName(fileName)];
  MailApp.sendEmail(defaultLogisticsEmail + "," + copyEmail,company + ' RMA as been fixed. (RMA:' + previusRmaId + ')', "The fixtures sent by " + company + " have been repaired and approved by the manager." + "\n" + "The attached file contain all the fixtures.", {attachments: blobs});

  Browser.msgBox("An email has been sent to the logistics and a copy email has been sent to you.");
}












//a function that being triggered every 10 minutes, going over the label that contains all the messages from customer (the messages got ther by gmail filter), every image attached to the meil get save to the folder
function receivedEmailProcessor()
{
  var finishedLable = GmailApp.getUserLabelByName("Processed costomer emails"); //the label the email will be moved after we got the image
  var lable = GmailApp.getUserLabelByName("Unprocessed customers emails"); //the lable of the email we need to get the images
  var emailsThreads = lable.getThreads(); //gets all the threads in the lable

  //going through all the threads in the lable
  for (var i = emailsThreads.length - 1; i >= 0;i--)
  {
    var messages = emailsThreads[i].getMessages(); //get the current thread massages
    var subject = messages[0].getSubject(); //get the subject of the current massage
    var rmaNumber = subject.slice(subject.lastIndexOf(':') + 2,); //extract the RMA number from the message subject
    var imageNumber = 1; //image attachment numer (in case there is more than one image attached)
    var firstImage = true;
    for (var j = 0; j < messages.length;j++) //going through the messages in the thread
    {
      var attachments = messages[j].getAttachments(); //getting the attachments of the current message
      for (var k = 0; k < attachments.length;k++) //going through all the attachments in the message
      {
        var attachmentType = attachments[k].getContentType();
        if (attachmentType == "image/jpeg" || attachmentType == "image/png") //checking if the current attachment is an image
        {
          var imageName = "RMA:" + rmaNumber + " delivery image(" + imageNumber + ")";

          if (firstImage == true) //if its the first image of this rma we will insert a link in the database and creat a new folder for the images.
          {
            firstImage = false;
            var newFolderId = DriveApp.getFolderById("1r8wKiLOidGlFCCPwYvynho8vgYkpGSGX").createFolder("RMA: " + rmaNumber).getId(); //Creating a new folder for this rma, where all the photos of this rma will be saved.
            var imageId = DriveApp.getFolderById(newFolderId).createFile(attachments[k]).getId(); //inserting the attached image to the folder
            DriveApp.getFileById(imageId).setName(imageName); //changing the name of the image
            var imageUrl = DriveApp.getFileById(imageId).getUrl();
            imageNumber++;

            //inserting the new image url to the database
            var databaseSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1B3PIk0T0F4txYTE9daP1pVCdhH2t_zVGd6Uz_SLuZi4/edit#gid=55465917").getSheetByName("CUSTOMERS RMA");

            var lastRow = databaseSheet.getLastRow(); //getting the last row with data.
            if (databaseSheet.getRange(lastRow,1).getValue() == "")
            {
              databaseSheet.getRange(lastRow,1).activate();
              databaseSheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
              lastRow = databaseSheet.getActiveCell().getRowIndex();
            }
          
            for (var currentRow = 1; currentRow <= lastRow; currentRow++)
            {
              if (databaseSheet.getRange(currentRow,1).getValue() == rmaNumber) //if it's the rma row, we insert the image link
              {
                databaseSheet.getRange(currentRow,21).setRichTextValue(SpreadsheetApp.newRichTextValue() 
        .setText(imageUrl)
        .setTextStyle(0, 82, SpreadsheetApp.newTextStyle()
        .setForegroundColor('#1155cc')
        .setUnderline(true)
        .build())
        .build());
                databaseSheet.getRange(currentRow,21).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
              }
            }
          }

          else //if it's not the first image of this rma
          {
            var imageId = DriveApp.getFolderById(newFolderId).createFile(attachments[k]).getId(); //adding the image to the folder
            DriveApp.getFileById(imageId).setName(imageName); //changing the name of the image
            imageNumber++;
          }
        }
      }
    }
      

      //moving the current thread to the processed emails lable
      emailsThreads[i].removeLabel(lable).refresh();
      emailsThreads[i].addLabel(finishedLable).refresh();
    

    
  }
}







//A function being triggered every 10 minutes and checks if a new RMA request as been recived from a customer.
function rmaNotifier()
{
  var thisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CUSTOMERS RMA");
  var i = 2;
  while (thisSheet.getRange(i,2).getValue() != "") //going through all the rows with data
  {
    if (thisSheet.getRange(i,24).getValue() != true) //if the current row is a new one
    {
      var company = thisSheet.getRange(i,9).getValue();
      var formSubmissionNumber = thisSheet.getRange(i,2).getValue();
      GmailApp.sendEmail('juganurmasystem@gmail.com','New RMA request from ' + company,"You have got a new RMA reqest from " + company + ".")
      while (thisSheet.getRange(i,2).getValue() == formSubmissionNumber) //going through all the rows with the same form submission number and mark as not new
      {
        thisSheet.getRange(i,24).setValue(true);
        i++;
      }
      return;
    }
    i++;
  }
}







//an automatic email being send to the customer to ship the fixtures
function shipmentEmail()
{
  var activeRangeRow = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().getRow();
  var company = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(activeRangeRow,9).getValue();
  var customerDefaultEmail = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(activeRangeRow,16).getValue();
  
  var inputCustomerEmail = Browser.inputBox("The customer default Email to be sent is: " + customerDefaultEmail + "\\n" + "If you want the Email to be sent to another address, put the address in the input box. otherwise, leave it blank.");
  if (inputCustomerEmail == "cancel") {return;} //if the user click the X of the message box
  else if (inputCustomerEmail == "")
    var customerEmail = customerDefaultEmail;
  else
    var customerEmail = inputCustomerEmail;

  var emailBody = "Dear " + company + ", \n" + "you are requested to ship the RMA fixtures for repairing." + "\n" + "Please be in contact with Juganu's Global Logistics Manager, Mrs. Hana Krichli - Hanak@juganu.com";

  GmailApp.sendEmail(customerEmail,"RMA shipment " + company,emailBody,{htmlBody: emailBody,
  cc: 'Adik@juganu.com' });

  Browser.msgBox("An email as been sent to the customer")

 
}

