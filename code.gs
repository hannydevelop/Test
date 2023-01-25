const DEFAULT_INPUT_TEXT = '';
const DEFAULT_OUTPUT_TEXT = '';

var newSheetSection = CardService.newCardSection();
var inputSheetSection = CardService.newCardSection();
var buttonSheetSection = CardService.newCardSection();

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
    if( row[3] == true) {
      var emailAddress = row[0];  // First column
      var message = row[1];       // Second column
      var subject = "Task Item Due";
      try {
        Logger.log('emailAddress')
          MailApp.sendEmail(emailAddress, subject, message);
          Logger.log(emailAddress)
      } catch(errorDetails) {
        Logger.log(errorDetails);
        // MailApp.sendEmail("eddyparkinson@someaddress.com", "sendEmail script error", errorDetails.message);
      }

    }
  }

}

function onHomepage() {
  sendEmails();
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

  var optionalArgs = {valueInputOption: "USER_ENTERED"};
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
