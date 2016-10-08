var SPREADSHEET_MAIN = 'Operaciones';
var SPREADSHEET_MASTER = 'Master';
var spreadsheet;
var masterSheet;
var defaults = getDefaults();

setSheets();

var ROW_START = 2;
var OPERATIONS_RANGE = operationsSheet.getRange('D2:D200');
var AMOUNT_RANGE = operationsSheet.getRange('F2:F200');
var DETAIL_RANGE = operationsSheet.getRange('H2:H200');

var OPERATION_TYPES = [
  'ABONO',
  'COMISION',
  'PAGOS',
  'FONDEO',
  'RETIRO'
];

function setSheets() {
    operationsSheet = SpreadsheetApp.getActiveSpreadsheet();
    masterSheet = spreadsheet.getSheetByName(SPREADSHEET_MASTER);
}


function onOpen() {
 setSheets();
  var menuEntries = [
    {name: "Process Operations", functionName: "processOperations"}
  ];
  spreadsheet.addMenu("Stocks To Ticks", menuEntries);
}

function processOperations() {

}

function processStocks() {
    var stockName = '';
    var tick;
    var purchasePrice;

    for (var i = ROW_START, operation; (operation = operationsSheet.getRange('D' + i)); i++) {
        operationType = operation.getValue();

    }

    for (var i = 2, stock; (stock = masterSheet.getRange('B' + i)); i++) {
        stockName = stock.getValue();
        if (stockName !== '') {
            tick = defaults[stockName];
            masterSheet.getRange('F' + i).setValue(tick);
            purchasePrice = masterSheet.getRange('D' + i).getValue() / masterSheet.getRange('E' + i).getValue();
            masterSheet.getRange('G' + i).setValue(purchasePrice);
            masterSheet.getRange('H' + i).setValue('Buy');
        } else {
            break;
        }
    }
}