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

var TOTAL_MAPPINGS = {
    'ABONO': 'B1',
    'COMISION': 'B6',
    'PAGOS': 'B8',
    'FONDEO': 'B4',
    'RETIRO': 'B2'
};

var PAYMENT_MAPPINGS = {
    'Principal': 'B9',
    'Interes': 'B10',
    'ImpuestoInteres': 'B11',
    'Moratorios': 'B12',
    'ImpuestoMoratorios': 'B13'
};

function setSheets() {
    operationsSheet = SpreadsheetApp.getActiveSpreadsheet();
}


function onOpen() {
 setSheets();
  var menuEntries = [
    {name: "Process Operations", functionName: "processOperations"}
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

function processOperations() {
    var operationType,
        totalsByType,
        amount;

    totalsByType = initTotalsByType(OPERATION_TYPES);

    for (var i = ROW_START, operation; (operation = operationsSheet.getRange('D' + i)); i++) {
        operationType = operation.getValue();
        if (operationType !== '') {
            if (totalsByType[operationType] !== null) {
                // Get amount
                amount = operationsSheet.getRange('F' + i).getValue();
                totalsByType[operationType] += amount;
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

    processPayments();
}

// Principal: 151.2554 Interes: 33.8447 Impuesto Interes: 5.41515 Moratorios: 0 Impuesto Moratorios: 0
function processPayments() {
    var totalsByType,
        detail;

    totalsByType = initTotalsByType(PAYMENT_TYPES);

    for (var i = ROW_START, operation; (operation = operationsSheet.getRange('D' + i)); i++) {
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
                            totalsByType[paymentData[0]] += Math.abs(paymentData[1]);
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
        var range = PAYMENT_MAPPINGS[paymentType];
        operationsSheet.getRange(range).setValue(total);
    }
}