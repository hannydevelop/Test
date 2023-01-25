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

function onHomepage() {
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
      .setOnClickAction(CardService.newAction().setFunctionName('translateText'))
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
