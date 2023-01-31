const DEFAULT_INPUT_TEXT = '';
const DEFAULT_OUTPUT_TEXT = '';

var newSheetSection = CardService.newCardSection();
var inputSheetSection = CardService.newCardSection();
var buttonSheetSection = CardService.newCardSection();
var invoiceSection = CardService.newCardSection();
var navigationSection = CardService.newCardSection();
var sendInvoiceSection = CardService.newCardSection();

const INPUT_MAP = [
  { text: 'Bank', val: 'Bank' },
  { text: 'Cash', val: 'Cash' },
  { text: 'Loan', val: 'Loan' },
  { text: 'Credit card', val: 'Credit card' },
  { text: 'Sales', val: 'Sales' },
  { text: 'Services', val: 'Services' },
]

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  /// e.g.  var spreadsheet = SpreadsheetApp.openById('0AkGlO9jJLGO8dDJad3VNTkhJcHR3UXlJSVRNTFJreWc');     

  var range = sheet.getRange("B1");
  var dateString = new Date().toString();
  range.setValue(dateString);   // this makes all formulas recalculate

  var startRow = 3;  // First row of data to process
  var numRows = 999;   // Number of rows to process
  // Fetch the range of cells
  var dataRange = sheet.getRange(startRow, 3, numRows, 7)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  var date = new Date().getDate().toString();

  for (i in data) {
    var row = data[i];
    if (row[4] == 'no' && date) {
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
    "left_margin=0.0&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + gid + '&';

  var params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };

  var response = UrlFetchApp.fetch(url, params).getBlob();
  // save to drive
  //  DriveApp.createFile(response);


  //or send as email
  var email = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B13").getValue();
  var company = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B2").getValue();
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
  var date = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F10").getDisplayValue();
  var invoiceNum = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F9").getValue();
  var description = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B16:C16").getValue();
  var dueDate = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F11").getDisplayValue();
  var amount = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F31").getValue();
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

// update transaction if invoice is completed
function blah() {
  var spreadsheet = SpreadsheetApp.openById('1_f0kSp0jD5LLG_HRfzJymq-YDQc52qw58F9CEMBV7ps');
  /// e.g.  var spreadsheet = SpreadsheetApp.openById('0AkGlO9jJLGO8dDJad3VNTkhJcHR3UXlJSVRNTFJreWc');     

  var sheet = spreadsheet.getSheets()[2];

  var dataRange = sheet.getRange(1, 7, 999);
  var values = dataRange.getValues();



  var optionalArgs = { valueInputOption: "USER_ENTERED" };

  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == "yes") {
        var request = {
          "majorDimension": "ROWS",
          "values": [
            [
              values[row][col]
            ]
          ]
        };
        Sheets.Spreadsheets.Values.append(
          request,
          spreadsheetId,
          'Transactions!A:E',
          optionalArgs
        )
      }
    }
  }
}

function onHomepage() {
  ScriptApp.newTrigger('sendEmails')
    .timeBased()
    .atHour(7)
    .nearMinute(0)
    .everyDays(1)
    .create();
}

function onDrive() {
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

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Manage all bookkeeping in one place. Start by creating a Spreadsheet"))
    .addSection(newSheetSection)
    .build();
  return card;
}

function invoice() {
  // create invoice
  var contactName = CardService.newTextInput()
    .setFieldName(`Contact Name`)
    .setTitle(`Receiver's name`);

  var clientName = CardService.newTextInput()
    .setFieldName(`Client Company`)
    .setTitle(`Client Company's Name`);

  var clientAddress = CardService.newTextInput()
    .setFieldName(`Client Address`)
    .setTitle(`Client Company's Address`);

  var dueDate = CardService.newDatePicker()
    .setFieldName('Due Date')
    .setTitle('Due Date')

  var paymentTerms = CardService.newTextInput()
    .setFieldName(`PayTerms`)
    .setTitle(`Warranty, returns policy...`);

  var totalTax = CardService.newTextInput()
    .setFieldName(`Tax`)
    .setTitle(`Total Tax (Optional)`);

  var discount = CardService.newTextInput()
    .setFieldName(`Discount`)
    .setTitle(`Discount (Optional)`);

  var email = CardService.newTextInput()
    .setFieldName(`Client Email`)
    .setTitle(`Client Email`);

  var postInvoice = CardService.newAction()
    .setFunctionName('postInvoice');
  var newpostInvoiceButton = CardService.newTextButton()
    .setText('View Invoice')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(postInvoice);

  // send invoice to user.
  var invoiceName = CardService.newTextInput()
    .setFieldName('Invoice Name')
    .setTitle('Invoice Name');
  var sendInvoice = CardService.newAction()
    .setFunctionName('sendInvoice');
  var newInvoiceButton = CardService.newTextButton()
    .setText('Send Invoice')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(sendInvoice);

  invoiceSection.addWidget(contactName);
  invoiceSection.addWidget(clientName);
  invoiceSection.addWidget(email);
  invoiceSection.addWidget(clientAddress);
  invoiceSection.addWidget(dueDate);
  invoiceSection.addWidget(discount);
  invoiceSection.addWidget(totalTax);
  invoiceSection.addWidget(paymentTerms);
  invoiceSection.addWidget(CardService.newButtonSet().addButton(newpostInvoiceButton));

  sendInvoiceSection.addWidget(invoiceName);
  sendInvoiceSection.addWidget(CardService.newButtonSet().addButton(newInvoiceButton));

  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Create, Send and Track Invoices"))
    .addSection(invoiceSection)
    .addSection(sendInvoiceSection)
    .build();
  return card;
}

function transaction() {
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
    .setHeader(CardService.newCardHeader().setTitle("Record Transactions"))
    .addSection(inputSheetSection)
    .addSection(buttonSheetSection)
    .build();
  return card;
}

function onSheet() {
  var buttonAction = CardService.newAction()
    .setFunctionName('invoice');
  navigationSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Record Transactions in Sheet")
    .setIconUrl('https://www.linkpicture.com/q/book_5.png')
    .setEndIcon(CardService.newIconImage().setIconUrl('https://www.linkpicture.com/q/icons8-forward-button-64.png'))
    .setText('Invoice Actions')
    .setOnClickAction(buttonAction));

  var buttonAction = CardService.newAction()
    .setFunctionName('transaction');
  navigationSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Create, Send and Track Invoices")
    .setIconUrl('https://www.linkpicture.com/q/bookkeeping.png')
    .setEndIcon(CardService.newIconImage().setIconUrl('https://www.linkpicture.com/q/icons8-forward-button-64.png'))
    .setText('Transaction Actions')
    .setOnClickAction(buttonAction));

  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Perform all bookkeeping actions in your sheet").setImageUrl('https://www.linkpicture.com/q/IMG_2430.png'))
    .addSection(navigationSection)
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

function postInvoice(e) {
  var res = e['formInput'];
  var contactName = res['Contact Name'] ? res['Contact Name'] : '';
  var clientName = res['Client Company'] ? res['Client Company'] : '';
  var clientEmail = res['Client Email'] ? res['Client Email'] : '';
  var clientAddress = res['Client Address'] ? res['Client Address'] : '';
  var dueDate = res['Due Date'] ? res['Due Date'] : '';
  var payTerms = res['PayTerms'] ? res['PayTerms'] : '';
  var discount = res['Discount'] ? res['Discount'] : 0;
  var totalTax = res['totalTax'] ? res['totalTax'] : 0;
  const invNumber = Math.floor(100000 + Math.random() * 900000);

  let date = dueDate.msSinceEpoch;
  // WE NEED TO RETRIEVE USER'S TIMEZONE
  let formatDate = Utilities.formatDate(new Date(date), "GMT", "yyyy/MM/dd")

  // set today's date as invoice.
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;

  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B10").setValue(contactName);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B11").setValue(clientName);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B12").setValue(clientAddress);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B13").setValue(clientEmail);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F11").setValue(formatDate);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F26").setValue(discount);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F28").setValue(totalTax);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F9").setValue(`INV${invNumber}`);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B35:F35").setValue(payTerms);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F10").setValue(formattedToday);
}

function submitRecord(e) {
  var res = e['formInput'];


  var Description = res['Description'] ? res['Description'] : 'Expenses';
  var Amount = res['Amount'] ? res['Amount'] : 'Expenses';
  var Debit = res['Debit'] ? res['Debit'] : 'Expenses';
  var Credit = res['Credit'] ? res['Credit'] : 'Expenses';


  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;
  const transactiomNumber = Math.floor(100000 + Math.random() * 900000);

  // Add today's date
  // Add unique reference number

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
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
    'Transactions!A:E',
    optionalArgs
  )

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Recorded Transaction`))
    .build();

}
