var SPREADSHEET_MAIN = 'Operaciones';
var SPREADSHEET_MASTER = 'Master';
var spreadsheet;
var masterSheet;

setSheets();
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
  'ImpuestoMoratorios',
  'IVAComisionMoratorios'
];

var PAYMENT_TYPES_COLUMNS = [
    'L', 'M', 'N', 'O', 'P','Q'
];

function setSheets() {
    operationsSheet = SpreadsheetApp.getActiveSpreadsheet();
}


function onOpen() {
 setSheets();
  var menuEntries = [
    {name: "Process Selected Payments", functionName: "processSelectedPayments"},
    {name: "Process Date", functionName: "processDate"}
  ];
  operationsSheet.addMenu("Process Operations", menuEntries);
}

// Principal: 151.2554 Interes: 33.8447 Impuesto Interes: 5.41515 Moratorios: 0 Impuesto Moratorios: 0
function processSelectedPayments() {
    var rowStart;

    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Start Row Number', 'What is first row number you want to process?', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK) {
        rowStart = Math.abs(response.getResponseText());
    } else if (response.getSelectedButton() == ui.Button.CANCEL) {
        Logger.log('The user canceled the dialog.');
    } else {
        Logger.log('The user closed the dialog.');
    }
    processPaymentRows(rowStart);
}

function processDate() {
    var rowStart;

    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Start Row Number', 'What is first row number you want to process?', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK) {
        rowStart = Math.abs(response.getResponseText());
    } else if (response.getSelectedButton() == ui.Button.CANCEL) {
        Logger.log('The user canceled the dialog.');
    } else {
        Logger.log('The user closed the dialog.');
    }

    for (var i = rowStart, operation; (operation = operationsSheet.getRange('D' + i)); i++) {
        operationType = operation.getValue();

        if (operationType !== '') {
             // Set date
            date = operationsSheet.getRange('B' + i).getValue();
            day = date.substring(0,2);
            month = date.substring(3,5);
            year = date.substring(6,10);
            formattedDate = year + '/' + month + '/' + day;
            operationsSheet.getRange('K' + i).setValue(formattedDate);
        } else {
            break;
        }
    }
}



function processPaymentRows(rowStart) {
    var date,
        year,
        month,
        day,
        formattedDate,
        paymentsByType,
        paymentTypeColumn,
        paymentAmount,
        detail;

    for (var i = rowStart, operation; (operation = operationsSheet.getRange('D' + i)); i++) {
        operationType = operation.getValue();

        if (operationType !== '') {
             // Set date
            date = operationsSheet.getRange('B' + i).getValue();
            date = date + '';
            day = date.substring(0,2);
            month = date.substring(3,5);
            year = date.substring(6,10);
            formattedDate = year + '/' + month + '/' + day;
            operationsSheet.getRange('K' + i).setValue(formattedDate);

            // Split then process
            if (operationType !== 'PAGOS') {
                continue;
            };

            detail = operationsSheet.getRange('H' + i).getValue();
            if (detail !== '') {
                var replaceStr = detail.replace(/: /gi, ':');
                replaceStr = replaceStr.replace(/Impuesto /gi, 'Impuesto');
                replaceStr = replaceStr.replace(/IVA Comisión Moratorios/gi, 'IVAComisionMoratorios');
                var payments = replaceStr.split(' ');
                for (var j = 0, payment; (payment = payments[j]); j++) {
                    if (payment !== null && payment !== '') {
                        var paymentData = payment.split(':');
                        if (paymentData.length === 2) {
                            paymentAmount = Math.abs(paymentData[1]);
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
}