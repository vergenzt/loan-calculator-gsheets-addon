/**
 * @OnlyCurrentDoc Limits the script to only accessing the current spreadsheet.
 */

function onInstall() {
  onOpen();
}
function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Use in this spreadsheet', 'use')
    .addToUi();
}
function use() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Loan Calculation Functions',
    'Custom functions are now available.',
    ui.ButtonSet.OK
  );
}

/**
 * @customfunction
 */
function ACCRUE(loans, accounts, dateUpdated, payments, paymentDates, asOf) {
  loans = headerize(loans);
  accounts = headerize(accounts);
  if (!(payments instanceof Array)) payments = [[payments]];
  if (!(paymentDates instanceof Array)) paymentDates = [[paymentDates]];

  var curDate = dateUpdated;

  while (payments.length > 0 && (!asOf || curDate <= asOf)) {
    // increment date
    curDate.setDate(curDate.getDate() + 1);

    // if payment is happening on this date, do it
    if (curDate >= paymentDates[0][0]) {
      makePayment(payments[0][0]);
      payments.shift();
      paymentDates.shift();
    }

    // increment interest for each loan
    for (var i=1; i<loans.length; i++) {
      if (loans[i].InterestDate <= curDate) {
        incrementInterest(loans[i]);
      }
    }
  }

  function makePayment(paymentAmt, toAccount) {
    // if no account is specified, then first make minimum payments on all accounts first
    if (!toAccount) {
      for (var i=1; i < accounts.length; i++) {
        var dueDate = accounts[i].MinPmtDate;
        if (dueDate <= curDate) {
          makePayment(accounts[i].MinPmt, accounts[i].Account);
          dueDate.setMonth(dueDate.getMonth() + 1);
        }
      }
    }

    // pay off by order in table
    for (var i=1; i<loans.length; i++) {
      var loan = loans[i];

      // skip if not toAccount
      if (toAccount && loan.Account != toAccount)
        continue;

      // pay interest first, then principal
      ['Interest', 'Principal']
        .forEach(function (field) {
          if (loan[field] > 0) {
            var amount = Math.min(paymentAmt, loan[field]);
            loan[field] -= amount;
            paymentAmt  -= amount;
          }
        });

      // restore balance
      loan.Balance = loan.Principal + loan.Interest;
    }
  }

  function incrementInterest(loan) {
    if (loan.InterestDate <= curDate) {
      loan.Interest = loan.Interest + loan.Balance * (Math.pow(1 + loan.Rate, 1/365) - 1);
      loan.Balance = loan.Principal + loan.Interest;
    }
  }

  loans = unheaderize(loans);
  return loans;
}


/**
 * Return total Balance, Principal, and Interest of loans.
 * @customFunction
 */
function SUMMARIZE(loans) {
  var results = [0, 0, 0, 0];
  headerize(loans).forEach(function(loan, i) {
    if (i == 0) return;
    results[0] += loan.Balance;
    results[1] += loan.Principal;
    results[2] += loan.Interest;
    if (loan.Balance > 0) results[3]++;
  });
  return results;
};


/*******************
 * Utility methods *
 *******************/

function headerize(table) {
  var header = table[0];
  return table.map(function(loan, i) {
    if (i == 0) return loan;
    else {
      var obj = {};
      for (var j=0; j<header.length; j++) {
        obj[header[j]] = loan[j];
      }
      return obj;
    }
  });
}

function unheaderize(table) {
  var header = table[0];
  return table.map(function(loan, i) {
    if (i == 0) return loan;
    else {
      var arr = [];
      for (var j=0; j<header.length; j++) {
        arr[j] = loan[header[j]];
      }
      return arr;
    }
  });
}
