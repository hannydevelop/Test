const DEFAULT_INPUT_TEXT = '';
const DEFAULT_OUTPUT_TEXT = '';

var newSheetSection = CardService.newCardSection();
var inputSheetSection = CardService.newCardSection();
var buttonSheetSection = CardService.newCardSection();
var invoiceSection = CardService.newCardSection();

const INPUT_MAP = [
  { text: 'Bank', val: 'Bank' },
  { text: 'Cash', val: 'Cash' },
  { text: 'Loan', val: 'Loan' },
  { text: 'Credit card', val: 'Credit card' },
  { text: 'Sales', val: 'Sales' },
  { text: 'Services', val: 'Services' },
]

function sendEmails() {
  var spreadsheet = SpreadsheetApp.openById('1y19ymdqET1R5uVk5uhs9PfbReOVkYUIpR9frAQjN3qM');
  /// e.g.  var spreadsheet = SpreadsheetApp.openById('0AkGlO9jJLGO8dDJad3VNTkhJcHR3UXlJSVRNTFJreWc');     

  var sheet = spreadsheet.getSheets()[0]; // gets the first sheet, i.e. sheet 0

  var range = sheet.getRange("B1");
  var dateString = new Date().toString();
  range.setValue(dateString);   // this makes all formulas recalculate

  var startRow = 4;  // First row of data to process
  var numRows = 50;   // Number of rows to process
  // Fetch the range of cells
  var dataRange = sheet.getRange(startRow, 1, numRows, 4)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  for (i in data) {
    var row = data[i];
    if (row[3] == true) {
      var emailAddress = row[0];  // First column
      var message = row[1];       // Second column
      var subject = "Task Item Due";
      try {
        Logger.log('emailAddress')
        MailApp.sendEmail(emailAddress, subject, message);
        Logger.log(emailAddress)
      } catch (errorDetails) {
        Logger.log(errorDetails);
        // MailApp.sendEmail("eddyparkinson@someaddress.com", "sendEmail script error", errorDetails.message);
      }

    }
  }

}

function sendInvoice(e) {
  var res = e['formInput'];
  var invoiceName = res['Invoice Name'] ? res['Invoice Name'] : 'Invoice';

  const fr = 0, fc = 0, lc = 9, lr = 27;
  var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();
  var gid = SpreadsheetApp.getActiveSpreadsheet().getSheetId();
  var url = "https://docs.google.com/spreadsheets/d/" + ssID + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + gid + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  var params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };

  var response = UrlFetchApp.fetch(url, params).getBlob();
  // save to drive
  //  DriveApp.createFile(response);


  //or send as email
  var email = SpreadsheetApp.getActiveSheet().getRange("invoicegen!B13").getValue();
  var company = SpreadsheetApp.getActiveSheet().getRange("invoicegen!B2").getValue();
  var subject = `Invoice From ${company}`;
  var body = 'Invoice Ready';

  MailApp.sendEmail(email, subject, body, {
    attachments: [{
      fileName: invoiceName + ".pdf",
      content: response.getBytes(),
      mimeType: "application/pdf"
    }]
  })

    // send info to invoice template record.
  var date = SpreadsheetApp.getActiveSheet().getRange("invoicegen!F10").getValue();
  var invoiceNum = SpreadsheetApp.getActiveSheet().getRange("invoicegen!F9").getValue();
  var description = SpreadsheetApp.getActiveSheet().getRange("invoicegen!B16:C16").getValue();
  var dueDate = SpreadsheetApp.getActiveSheet().getRange("invoicegen!F11").getValue(); 
  var amount = SpreadsheetApp.getActiveSheet().getRange("invoicegen!F31").getValue(); 
  // var receiverAddr = email
  
    var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        date,
        invoiceNum,
        description,
        dueDate,
        amount,
        email
      ]
    ]
  }

  var optionalArgs = { valueInputOption: "USER_ENTERED" };
  Sheets.Spreadsheets.Values.append(
    request,
    ssID,
    'invoice!A:E',
    optionalArgs
  )


  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly sent invoice to ${email}`))
    .build();
}

function onHomepage() {
  // Trigger to send email reminder everyday at 7:30am in their time zone.

  /*ScriptApp.newTrigger('sendEmails')
    .timeBased()
    .atHour(1)
    .nearMinute(7)
    .everyDays(1)
    .create();
  */

  var invoiceName = CardService.newTextInput()
    .setFieldName('Invoice Name')
    .setTitle('Invoice Name');
  var sendInvoice = CardService.newAction()
    .setFunctionName('sendInvoice');
  var newInvoiceButton = CardService.newTextButton()
    .setText('Send Invoice')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(sendInvoice);
    invoiceSection.addWidget(invoiceName);
  invoiceSection.addWidget(CardService.newButtonSet().addButton(newInvoiceButton));

  var sheetName = CardService.newTextInput()
    .setFieldName('Sheet Name')
    .setTitle('Sheet Name');
  var createNewSheet = CardService.newAction()
    .setFunctionName('copyFile');
  var newSheetButton = CardService.newTextButton()
    .setText('Create New Sheet')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(createNewSheet);
  newSheetSection.addWidget(sheetName);
  newSheetSection.addWidget(CardService.newButtonSet().addButton(newSheetButton));

  var description = CardService.newTextInput()
    .setFieldName('Description')
    .setTitle('Description');

  var amount = CardService.newTextInput()
    .setFieldName('Amount')
    .setTitle('Amount');

  var debit = CardService.newSelectionInput().setTitle('From')
    .setFieldName('Debit')
    .setType(CardService.SelectionInputType.DROPDOWN);

  INPUT_MAP.forEach((language, index, array) => {
    debit.addItem(language.text, language.val, language.val == true);
  })

  var credit = CardService.newSelectionInput().setTitle('To')
    .setFieldName('Credit')
    .setType(CardService.SelectionInputType.DROPDOWN);

  INPUT_MAP.forEach((language, index, array) => {
    credit.addItem(language.text, language.val, language.val == true);
  })

  inputSheetSection.addWidget(description);
  inputSheetSection.addWidget(amount);
  inputSheetSection.addWidget(debit);
  inputSheetSection.addWidget(credit);


  buttonSheetSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Record Transaction')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('submitRecord'))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Clear')
      .setOnClickAction(CardService.newAction().setFunctionName('clearText'))
      .setDisabled(false)));

  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Card title"))
    .addSection(newSheetSection)
    .addSection(inputSheetSection)
    .addSection(buttonSheetSection)
    .addSection(invoiceSection)
    .build();
  return card;
}

function copyFile(e) {
  var res = e['formInput'];
  var sheetName = res['Sheet Name'] ? res['Sheet Name'] : 'Expenses';
  let file = DriveApp.getFileById('1S4GMiZ0H0_6OHH7DEnjZt07-6kk0eMP4YSNUmRcKZXA');
  file = file.makeCopy().setName(sheetName);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly created the file ${sheetName}`))
    .build();
}

function submitRecord(e) {
  var res = e['formInput'];


  var Description = res['Description'] ? res['Description'] : 'Expenses';
  var Amount = res['Amount'] ? res['Amount'] : 'Expenses';
  var Debit = res['Debit'] ? res['Debit'] : 'Expenses';
  var Credit = res['Credit'] ? res['Credit'] : 'Expenses';


  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // Add today's date
  // Add unique reference number

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        '2023-01-01',
        '1',
        Description,
        Amount,
        Debit,
        Credit
      ]
    ]
  }

  var optionalArgs = { valueInputOption: "USER_ENTERED" };
  Sheets.Spreadsheets.Values.append(
    request,
    spreadsheetId,
    'A:E',
    optionalArgs
  )

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Recorded Transaction`))
    .build();

}
