var SPREADSHEET_MAIN = 'Operaciones';
var SPREADSHEET_MASTER = 'Master';
var spreadsheet;
var masterSheet;

setSheets();

var ROW_START = 17;
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

var PAYMENT_TYPES = [
  'Principal',
  'Interes',
  'ImpuestoInteres',
  'Moratorios',
  'ImpuestoMoratorios'
];

var PAYMENT_TYPES_COLUMNS = [
    'L', 'M', 'N', 'O', 'P'
];

var TOTAL_MAPPINGS = {
    'ABONO': 'B1',
    'COMISION': 'B6',
    'PAGOS': 'B8',
    'FONDEO': 'B4',
    'RETIRO': 'B2'
};

var PAYMENT_MAPPINGS = {
    'Principal': 'B10',
    'Interes': 'B11',
    'ImpuestoInteres': 'B12',
    'Moratorios': 'B13',
    'ImpuestoMoratorios': 'B14'
};

var PAYMENT_MAPPINGS_ADDITIONAL = {
    'Principal': 'C10',
    'Interes': 'C11',
    'ImpuestoInteres': 'C12',
    'Moratorios': 'C13',
    'ImpuestoMoratorios': 'C14'
};

var PAYMENT_MAPPINGS_ADDITIONAL = {
    'Principal': 'D10',
    'Interes': 'D11',
    'ImpuestoInteres': 'D12',
    'Moratorios': 'D13',
    'ImpuestoMoratorios': 'D14'
};

function setSheets() {
    operationsSheet = SpreadsheetApp.getActiveSpreadsheet();
}


function onOpen() {
 setSheets();
  var menuEntries = [
    {name: "Process Operations", functionName: "processOperations"},
    {name: "Process All Payments", functionName: "processAllPayments"},
    {name: "Process Selected Payments", functionName: "processSelectedPayments"}
  ];
  operationsSheet.addMenu("Process Operations", menuEntries);
}

function initTotalsByType(types) {
    var totalsByType = {};
    for (var i = 0, type;(type = types[i]); i++) {
        totalsByType[type] = 0;
    }
    return totalsByType;
}

function initCurrentPayments() {
    var paymentsByType = {};
    var paymentValue;
    for (var i = 0, type;(type = PAYMENT_TYPES[i]); i++) {
        paymentValue = operationsSheet.getRange(PAYMENT_MAPPINGS[type]).getValue();
        paymentsByType[type] = paymentValue;
    }
    return paymentsByType;
}

function processOperations() {
    var operationType,
        totalsByType,
        date,
        amount;

    totalsByType = initTotalsByType(OPERATION_TYPES);

    for (var i = ROW_START, operation; (operation = operationsSheet.getRange('D' + i)); i++) {
        operationType = operation.getValue();
        if (operationType !== '') {
            if (totalsByType[operationType] !== null) {
                // Get amount
                amount = operationsSheet.getRange('F' + i).getValue();
                totalsByType[operationType] += amount;
                // Set date
                date = operationsSheet.getRange('B' + i).getValue();
                operationsSheet.getRange('K' + i).setValue(date.substring(0,10));
            }
        } else {
            break;
        }
    }

    for (var j = 0; (operationType = OPERATION_TYPES[j]); j++) {
        var total = totalsByType[operationType];
        var range = TOTAL_MAPPINGS[operationType];
        operationsSheet.getRange(range).setValue(total);
    }
}

// Principal: 151.2554 Interes: 33.8447 Impuesto Interes: 5.41515 Moratorios: 0 Impuesto Moratorios: 0
function processAllPayments() {
    var totalsByType;

    totalsByType = initTotalsByType(PAYMENT_TYPES);
    processPaymentRows(ROW_START, PAYMENT_MAPPINGS, totalsByType);
}

function processSelectedPayments() {
    var paymentsByType,
        rowStart,
        detail;

    paymentsByType = initCurrentPayments();

    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Start Row Number', 'What is first row number you want to process?', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK) {
        rowStart = response.getResponseText() + 0;
    } else if (response.getSelectedButton() == ui.Button.CANCEL) {
        Logger.log('The user canceled the dialog.');
    } else {
        Logger.log('The user closed the dialog.');
    }

    //TODO Process date

    processPaymentRows(rowStart, PAYMENT_MAPPINGS_ADDITIONAL, paymentsByType);
}



function processPaymentRows(rowStart, paymentMappings, totalsByType) {
    //TODO Remove sums
    var paymentsByType,
        paymentTypeColumn,
        paymentAmount,
        detail;

    for (var i = rowStart, operation; (operation = operationsSheet.getRange('D' + i)); i++) {
        operationType = operation.getValue();

        if (operationType !== '') {
            // Split then process
            if (operationType !== 'PAGOS') {
                continue;
            };

            detail = operationsSheet.getRange('H' + i).getValue();
            if (detail !== '') {
                var replaceStr = detail.replace(/: /gi, ':');
                replaceStr = replaceStr.replace(/Impuesto /gi, 'Impuesto');
                var payments = replaceStr.split(' ');
                for (var j = 0, payment; (payment = payments[j]); j++) {
                    if (payment !== null && payment !== '') {
                        var paymentData = payment.split(':');
                        if (paymentData.length === 2) {
                            paymentAmount = Math.abs(paymentData[1]);
                            totalsByType[paymentData[0]] += paymentAmount;
                            paymentTypeColumn = PAYMENT_TYPES_COLUMNS[j];
                            operationsSheet.getRange(paymentTypeColumn + i).setValue(paymentAmount);
                        }
                    }
                }
            }
        } else {
            break;
        }
    }

    for (var i = 0, paymentType; (paymentType = PAYMENT_TYPES[i]); i++) {
        var total = totalsByType[paymentType];
        var range = paymentMappings[paymentType];
        operationsSheet.getRange(range).setValue(total);
    }
}