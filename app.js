/*####################################
Tool Built by krishnan.ubuntu@gmail.com
######################################*/

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var menuEntries = [ {name: "Report an issue", functionName: "reportIssue"}];
  
 var menuEntries1 = [
                      {name: "Send Issue Status", functionName: "sendStatusToUser"},
                      {name: "Generate Report", functionName: "generateReport"} ];
  
  ss.addMenu("IT Helpdesk Menu", menuEntries);
  ss.addMenu("Admin Management Tools", menuEntries1);
}


function reportIssue (){
  var ss = SpreadsheetApp.openById("SPREADSHEET UNIQUE CODE");
  var listSheet = ss.getSheetByName('List Items');
  var ticketSheet = ss.getSheetByName('Live Issues');
  var app = UiApp.createApplication().setTitle('Report an issue');
  var panel = app.createVerticalPanel().setId('panel');
    var dateLabel = app.createLabel('Date');
  var dateValue = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd hh:mm:ss");
  var date = app.createTextBox().setId('date').setName('date').setValue(dateValue).setEnabled(false);
  var problemLabel = app.createLabel('Problem Category:');
  
  // List box code starts here
  var problemCategory = app.createListBox().setWidth('150x').setName('problemCategory').setId('problemCategory');
  numItemList1 = listSheet.getLastRow();
  //get the item array
  list1ItemArray = listSheet.getRange(1,1,numItemList1,1).getValues();
  //Add items in list box
  for(var i=0; i<list1ItemArray.length; i++){
    //listBox1.addItem(String(number));
    problemCategory.addItem(String(list1ItemArray[i][0]))
  }
//List box code ends here
  var userEmailLabel = app.createLabel('Enter your email:')
  var userEmail = app.createTextBox().setId('userEmail').setName('userEmail');
  var problemDescLabel = app.createLabel('Detailed Description of the problem:');
  var problemDescription = app.createTextArea().setId('problemDescription').setName('problemDescription');
   var submitButton = app.createButton('Submit').setId('submitButton');
  var clearButton = app.createButton('Clear').setId('clearButton');
  var confirmLabel = app.createLabel('Your issue has been successfully reported. We will resolve it ASAP. Thank You!').setId('confirmLabel').setStyleAttribute('color', 'red').setVisible(false);
  var problemErrorLabel = app.createLabel().setId('problemErrorLabel').setVisible(false);
  var emailErrorLabel = app.createLabel().setId('emailErrorLabel').setVisible(false);
  var prodDescriptionErrorLabel = app.createLabel().setId('prodDescriptionErrorLabel').setVisible(false);
  //Code for grid setup starts
  var grid = app.createGrid(7, 2);
  grid.setWidget(0, 0, dateLabel);
  grid.setWidget(0, 1, date);
  grid.setWidget(1, 0, problemLabel);
  grid.setWidget(1, 1, problemCategory);
  grid.setWidget(2, 0, userEmailLabel);
  grid.setWidget(2, 1, userEmail);
  grid.setWidget(3, 0, problemDescLabel);
  grid.setWidget(3, 1, problemDescription);
  grid.setWidget(4, 0, submitButton);
  grid.setWidget(4, 1, clearButton);
  grid.setWidget(5, 0, confirmLabel);
  grid.setWidget(5, 1, problemErrorLabel);
  grid.setWidget(6, 0, emailErrorLabel);
  grid.setWidget(6, 1, prodDescriptionErrorLabel);
    //Code for grid setup ends
  
  var clickHandler = app.createServerHandler("respondAndSubmitIssue");
  submitButton.addClickHandler(clickHandler);
  clickHandler.addCallbackElement(grid);

  var clickHandler1 = app.createServerHandler("respondAndClearReportIssue");
  clearButton.addClickHandler(clickHandler1);
  clickHandler1.addCallbackElement(grid);

  app.add(grid);
  var doc = SpreadsheetApp.getActive();
  doc.show(app);
  
}

function respondAndSubmitIssue(e){
  var ss = SpreadsheetApp.openById("SPREADSHEET UNIQUE CODE");
  var listSheet = ss.getSheetByName('List Items');
  var ticketSheet = ss.getSheetByName('Live Issues');
  var app = UiApp.getActiveApplication();
  var dateValue = e.parameter.date;
  var problemTypeValue = e.parameter.problemCategory;
  var userEmailValue = e.parameter.userEmail;
  var problemDescriptionContent = e.parameter.problemDescription;
  var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;
  var itTeamEmail = "Enter your service desk or group email";
  if (problemTypeValue == "Choose a category"){
   app.getElementById('problemErrorLabel').setText('Please select a product category').setStyleAttribute('color', 'red').setVisible(true);
  }
  else {
    if (!userEmailValue){
      app.getElementById('emailErrorLabel').setText('Please enter email address').setStyleAttribute('color', 'red').setVisible(true); 
    }
    else{ 
      if (emailPattern.test(userEmailValue) == false){
        app.getElementById('emailErrorLabel').setText('Please enter proper email address').setStyleAttribute('color', 'red').setVisible(true);
        }
      else{
      if (!problemDescriptionContent){
        app.getElementById('prodDescriptionErrorLabel').setText('Please provide problem description.').setStyleAttribute('color', 'red').setVisible(true); 
      }
      else{
  var lastCell = ticketSheet.getRange(ticketSheet.getLastRow()+1,1, 1, 4 ).setValues([[dateValue,userEmailValue,problemTypeValue,problemDescriptionContent]]);
        app.getElementById('emailErrorLabel').setVisible(false);
        app.getElementById('prodDescriptionErrorLabel').setVisible(false);
        app.getElementById('confirmLabel').setVisible(true);
        MailApp.sendEmail(userEmailValue, 'Issue Submitted Successfully', 'Your issue is submitted with the IT department. We will resolve your issue ASAP. \nSnapshot of your problem: \n Problem Category - ' + problemTypeValue + '\nProblem Description - '+ problemDescriptionContent);
        MailApp.sendEmail(itTeamEmail, 'New Issue Submited by - ' + userEmailValue, 'A new issue has been reported by - ' + userEmailValue + '\nPlease look into it ASAP. \nSnapshot of the problem: \n Problem Category - ' + problemTypeValue + '\nProblem Description - '+ problemDescriptionContent );

      }
  }
  }
  }
    return app;
}

  
function respondAndClearReportIssue (e){
  var ss = SpreadsheetApp.openById("SPREADSHEET UNIQUE CODE");
  var listSheet = ss.getSheetByName('List Items');
  var ticketSheet = ss.getSheetByName('Live Issues');
  var app = UiApp.getActiveApplication();
  app.getElementById('userEmail').setValue('');
  app.getElementById('problemDescription').setValue('');
  app.getElementById('confirmLabel').setVisible(false);
  app.getElementById('problemErrorLabel').setVisible(false);
  app.getElementById('emailErrorLabel').setVisible(false);
  app.getElementById('prodDescriptionErrorLabel').setVisible(false);
  return app;
}



function generateReport(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName('Resolved Issues');
 var app = UiApp.createApplication().setTitle("Generate Resolved Issues Report");
 var panel = app.createVerticalPanel();
  var horizontalPanel = app.createHorizontalPanel();
  // var currentDateValue = new Date().getTime() ;
  var currentDateValue = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd");
  var currentDateLabel = app.createLabel().setText("Today's date is:").setId('currentDateLabel');
  var currentDate = app.createTextBox().setName('currentDate').setId('currentDate').setValue(currentDateValue).setEnabled(false);
  var generateReportButton = app.createButton('Generate Report').setId('generateReportButton');
  var reportStatusLabel = app.createLabel().setText('Report Generated Successfully').setStyleAttribute('color', 'red').setId('reportStatusLabel').setVisible(false);
  var reportRelatedWarning = app.createLabel().setText('Note: You can find the report as Google Spreadsheet in your document list. Please fint the URL of the document below.').setStyleAttribute('color', 'green').setId('reportRelatedWarning').setVisible(false);
  var reportURLDisplay = app.createTextBox().setName('reportURLDisplay').setId('reportURLDisplay').setVisible(false);
  var fileNameDisplay = app.createTextBox().setId('fileNameDisplay').setName('fileNameDisplay').setVisible(false);
  var closeButton = app.createButton('Close').setId('closeButton').setVisible(false);
  var fileNameLabel = app.createLabel('Report Name').setId('fileNameLabel').setVisible(false);
  var fileUrlLabel = app.createLabel('Report URL').setId('fileUrlLabel').setVisible(false);
  var clickHandler = app.createServerHandler("respondToGenerateReport");
  generateReportButton.addClickHandler(clickHandler);
  clickHandler.addCallbackElement(panel);
  var clickHandler1 = app.createServerHandler("respondAndClose");
  closeButton.addClickHandler(clickHandler1);
  clickHandler1.addCallbackElement(panel);
  panel.add(currentDateLabel);
  panel.add(currentDate);
  panel.add(generateReportButton);
  panel.add(reportStatusLabel);
  panel.add(reportRelatedWarning);
  panel.add(fileNameLabel);
  panel.add(fileNameDisplay);
  panel.add(fileUrlLabel);
  panel.add(reportURLDisplay);
  panel.add(closeButton);
  app.add(panel);
  app.add(horizontalPanel);
  var doc = SpreadsheetApp.getActive();
  doc.show(app);
}

  
function respondToGenerateReport(e){
  //var ss = SpreadsheetApp.getActiveSheet();
  var ss = SpreadsheetApp.openById("SPREADSHEET UNIQUE CODE");
  var balanceSheet = ss.getSheetByName('Live Issues');
  var app = UiApp.getActiveApplication();
  var curDateValue = e.parameter.currentDate;
  var ssNew = SpreadsheetApp.create("Live Issues Issues - "+curDateValue);
  var spreadSheetName = String("Live Issues - "+curDateValue);
  balanceSheet.copyTo(ssNew);
  var newSheet = ssNew.getSheets()[0];
  newSheet.setName("Issues Report");
  var reportUrl = ssNew.getUrl();
  app.getElementById('reportStatusLabel').setVisible(true);
  app.getElementById('reportRelatedWarning').setVisible(true);
  app.getElementById('fileNameLabel').setVisible(true);
  app.getElementById('fileNameDisplay').setValue(spreadSheetName).setVisible(true);
  app.getElementById('fileUrlLabel').setVisible(true);
  app.getElementById('reportURLDisplay').setValue(reportUrl).setVisible(true);
  app.getElementById('closeButton').setVisible(true);
  //Sharing the report with the team.
  ssNew.addCollaborators("ithelpdesk@digitactical.com", {editorAccess:true, emailInvitations:true});
  return app;
}


function sendStatusToUser(){
 var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRowIndex();
  var userEmail = sheet.getRange(row, getColIndexByName("Username")).getValue();
  var subject = "Helpdesk Ticket #" + row;
  var body = "We've updated the status of your ticket.\n\nStatus: " + sheet.getRange(row, getColIndexByName("Status")).getValue();
    
  MailApp.sendEmail(userEmail, subject, body, {name:"Help Desk"});  

}


function respondAndClose(e){
  var app = UiApp.getActiveApplication();
  app.close();
  return app;  
}

//Library function
function getColIndexByName(colName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var numColumns = sheet.getLastColumn();
  var row = sheet.getRange(1, 1, 1, numColumns).getValues();
  for (i in row[0]) {
    var name = row[0][i];
    if (name == colName) {
      return parseInt(i) + 1;
    }
  }
  return -1;
}