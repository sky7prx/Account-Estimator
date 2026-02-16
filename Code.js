function onOpen (e) {
  const props = PropertiesService.getDocumentProperties();
  const initialized = props.getProperty('init');
  if (initialized !== 'true') {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Sheet automation').addItem('Enable automation','addTriggers').addToUi();

    const lastNotify = JSON.parse(props.getProperty('lastNotify')) || 0;
    const oneDay = 24 * 60 * 60 * 1000;
    const time = new Date().getTime();
    if (lastNotify < time - oneDay) {
      ui.alert(`This spreadsheet has built-in automation that makes several features work better. If you would like to enable all of the features, go to the 'Sheet automation' menu and select 'Enable automation'. Follow the Google prompts to authorize the advanced features.`);
      props.setProperty('lastNotify',JSON.stringify(time));
    }
  }
}

function addTriggers () {
  ScriptApp.newTrigger('runOnEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  const time = new Date().getTime();
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('init','true');
  props.setProperty('simpleLog',JSON.stringify([time]));
  props.setProperty('installableLog',JSON.stringify([time]));

  SpreadsheetApp.getUi().createMenu('Sheet automation').addItem('Done','onOpen').addToUi();
}

function onEdit (e) { //Logs times of edits captured with the simple trigger so we can make sure the installable trigger is running
  const time = new Date().getTime();
  const logLength = 5;
  const props = PropertiesService.getDocumentProperties();
  const log = JSON.parse(props.getProperty('simpleLog')) || [];
  const log2 = JSON.parse(props.getProperty('installableLog')) || [];
  console.log('simple log:',log);
  console.log('installable log:',log2);
  const buffer = 2 * 60 * 1000; //Check to see if there's a matching entry within 2 minutes
  let match = false;
  for (let i = 0; i < log.length; i++) {
    for (let j = 0; j < log2.length; j++) {
      if (Math.abs(log[i] - log2[j]) <= buffer) {
        match = true;
        break;
      }
    }
    if (match) break;
  }

  log.push(time);
  while (log.length > logLength) log.shift();
  props.setProperty('simpleLog',JSON.stringify(log));

  if (!match) {
    props.deleteProperty('init');
    onOpen(e);
  }
  console.log('match found?',match);
}

function runOnEdit(e) {
  const time = new Date().getTime();
  SpreadsheetApp.flush();

  //Save a log of times this function is triggered so we can check that it's executing using a simple trigger
  const logLength = 5;
  const props = PropertiesService.getDocumentProperties();
  const log = JSON.parse(props.getProperty('installableLog')) || [];
  log.push(time);
  while (log.length > logLength) log.shift();
  props.setProperty('installableLog',JSON.stringify(log));

  const ss = SpreadsheetApp.getActive();
  const event = JSON.parse(JSON.stringify(e));
  console.log(event);

  const firstTransRow = 5;

  /** @type {SpreadsheetApp.Sheet} */
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const actionSheets = ['Paychecks','Transactions','Archived Transactions','Recurring Transactions'];

  if (!actionSheets.includes(sheetName)) return;

  if (sheetName === 'Paychecks') {
    if (event.range.columnStart !== 1 ||
        event.range.columnEnd !== 1 ||
        event.range.rowStart !== event.range.rowEnd ||
        (event.value !== 'TRUE' && event.value !== true)) return;

    const payData = sheet.getRange(event.range.rowStart,2,1,3).getValues()[0];

    const transSheet = ss.getSheetByName('Transactions');
    const transRangeVals = transSheet.getRange(firstTransRow,9,transSheet.getMaxRows(),6).getValues();
    const transRangeReduced = transRangeVals.filter(el => el.join('') !== '');
    const nextRow = transRangeReduced.length + firstTransRow;

    payData.push('Paycheck');
    payData.push('Paycheck');

    transSheet.getRange(nextRow,9,1,5).setValues([payData]);
    //e.range.setValue(false);
  }
  if (sheetName === 'Transactions') {
    const a1 = e.range.getA1Notation();
    if (a1 === 'K3') { //Advance late transactions
      if (!['To today','To tomorrow'].includes(e.value)) return;

      e.range.setValue('');
      advanceTransactions(e.value);
    }
    if (a1 === 'L3') { //Sort transactions
      if (e.value !== 'TRUE') return;

      e.range.setValue(false);
      sortTrans();
    }
    if (a1 === 'G3') { //Change selected to today
      if (e.value !== 'TRUE') return;

      e.range.setValue(false);
      
      const today = new Date(new Date().toLocaleDateString());
      const sheetVals = sheet.getDataRange().getValues();

      const firstIndex = sheetVals[3].indexOf('Date') - 1;
      const lastIndex = sheetVals[3].lastIndexOf('Date') - 1;
      sheetVals.forEach((v,i) => {
        if (v[firstIndex] === true || v[firstIndex] === 'TRUE')
          sheet.getRange(1 + i,1 + firstIndex,1,2).setValues([[false,today]]);
      });
      sheetVals.forEach((v,i) => {
        if (v[lastIndex] === true || v[lastIndex] === 'TRUE')
          sheet.getRange(1 + i,1 + lastIndex,1,2).setValues([[false,today]]);
      });
      
      //sortTrans();
    }
  }
  if (sheetName === 'Archived Transactions') {
    if (e.value !== 'TRUE') return;

    const a1 = e.range.getA1Notation();
    if (a1 === 'L3') {
      e.range.setValue(false);
      sortTrans(sheetName);
    }
  }
  if (sheetName === 'Recurring Transactions') {
    if (e.value !== 'TRUE') return;

    const a1 = e.range.getA1Notation();
    if (a1 === 'K5') {
      e.range.setValue(false);
      placeRecurring();
    }
    if (a1 === 'K10') {
      e.range.setValue(false);
      archive();
    }
  }
}

function placeRecurring() { //Place the list of planned recurring transactions onto the Transactions tab and also create a new sheet for the selected month.
  const firstTransRow = 5;
  const ss = SpreadsheetApp.getActive();
  const recurSheet = ss.getSheetById(942392736);
  const transSheet = ss.getSheetById(1732160294);

  const month = recurSheet.getRange('L2').getValue();
  const year = recurSheet.getRange('L3').getValue();

  const props = PropertiesService.getDocumentProperties();
  const completedMonths = JSON.parse(props.getProperty('months')) || [];
  console.log('Completed months:',completedMonths);
  
  if (completedMonths.includes(month + year)) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('The recurring transactions for this month have already been copied to the Transactions tab. If you continue, it may cause duplicate transactions to appear. Are you sure you want to continue?',ui.ButtonSet.YES_NO);
    if (response === ui.Button.NO) return console.log('user ended script execution');
  }

  const maxDates = {
    January: 31,
    February: year % 4 === 0 ? 29 : 28,
    March: 31,
    April: 30,
    May: 31,
    June: 30,
    July: 31,
    August: 31,
    September: 30,
    October: 31,
    November: 30,
    December: 31
  };

  const setDate = d => new Date(month + " " + d + ", " + year);
  const isWeekend = d => {
    const weekday = Utilities.formatDate(d,Session.getScriptTimeZone(),"EEEE");
    return weekday === 'Saturday' || weekday === 'Sunday';
  };

  const holidays = getHolidays(year);
  const oneDay = 24 * 60 * 60 * 1000;

  const expenses = recurSheet.getRange(`A${firstTransRow-1}:D`).getValues().filter(r => r.join("") !== "").sort((a,b) => a[0] - b[0]).map(r => {
    r[0] = r[0] > maxDates[month] ? setDate(maxDates[month]) : setDate(r[0]);
    while (holidays.includes(r[0].toLocaleDateString()) || isWeekend(r[0])) {
      r[0] = new Date(r[0].getTime() - oneDay);
    }
    r[1] = isNaN(r[1]) ? "" : r[1];
    r.splice(2,0,"");
    return r;
  });
  const income = recurSheet.getRange(`F${firstTransRow-1}:I`).getValues().filter(r => r.join("") !== "").sort((a,b) => a[0] - b[0]).map(r => {
    r[0] = r[0] > maxDates[month] ? setDate(maxDates[month]) : setDate(r[0]);
    while (holidays.includes(r[0].toLocaleDateString()) || isWeekend(r[0])) {
      r[0] = new Date(r[0].getTime() - oneDay);
    }
    r[1] = isNaN(r[1]) ? "" : r[1];
    r.splice(2,0,"");
    return r;
  });

  const transExp = transSheet.getRange(`B${firstTransRow}:F`).getValues();
  const transInc = transSheet.getRange(`I${firstTransRow}:M`).getValues();

  const nextRow = rng => rng.reduce((lastIndex,row,rowIndex) => {
    if (row.join('') !== '') return rowIndex + 1;
    return lastIndex;
  },0);

  const nextExpRow = nextRow(transExp) + firstTransRow;
  const nextIncRow = nextRow(transInc) + firstTransRow;

  if (!completedMonths.includes(month + year)) completedMonths.push(month + year);
  props.setProperty('months',JSON.stringify(completedMonths));

  transSheet.getRange(nextExpRow,2,expenses.length,expenses[0].length).setValues(expenses);
  transSheet.getRange(nextIncRow,9,income.length,income[0].length).setValues(income);

  transSheet.getRange(`B${firstTransRow}:F`).sort({column: 2, ascending: false});
  transSheet.getRange(`I${firstTransRow}:M`).sort({column: 9, ascending: false});

  const prevSheetName = Utilities.formatDate(new Date(setDate(1).getTime() - (24 * 60 * 60 * 1000)),Session.getScriptTimeZone(),"MMM yyyy");
  const nextSheetName = Utilities.formatDate(setDate(1),Session.getScriptTimeZone(),"MMM yyyy");

  let nextMonthSheet = ss.getSheetByName(nextSheetName);
  if (!nextMonthSheet) {
    const prevMonthSheet = ss.getSheetByName(prevSheetName);
    nextMonthSheet = prevMonthSheet.copyTo(ss).setName(nextSheetName);
    nextMonthSheet.getRange('C7').setValue(setDate(1));
    nextMonthSheet.getRange('O12').setValue("");
    nextMonthSheet.setActiveRange(nextMonthSheet.getRange('A1'));
    ss.moveActiveSheet(2);
  }
}

function archive () {
  const firstTransRow = 5;
  const firstArchiveRow = 5;
  const ss = SpreadsheetApp.getActive();
  const recurSheet = ss.getSheetById(942392736); //Recurring transactions sheet
  const transSheet = ss.getSheetById(1732160294); //Transactions sheet
  const archiveSheet = ss.getSheetById(1366470246); //Archived Transactions sheet

  const month = recurSheet.getRange('L7').getValue(); //Selected month on the recurring transactions sheet
  const year = recurSheet.getRange('L8').getValue(); //Entered year on the recurring transactions sheet

  const monthIndex = { //List of 0-index months
    January: 0,
    February: 1,
    March: 2,
    April: 3,
    May: 4,
    June: 5,
    July: 6,
    August: 7,
    September: 8,
    October: 9,
    November: 10,
    December: 11
  };

  const selectedMonth = new Date(year,monthIndex[month],1).getTime(); //Return the 1st day of the selected month in milisecond format
  const nextMonth = new Date(year, monthIndex[month] + 1,1).getTime();
  const oneDay = 24 * 60 * 60 * 1000; //Calculate the number of miliseconds in one day
  const endOfMonth = new Date(nextMonth - oneDay).getTime(); //Return the last day of the selected month in milisecond format

  const archiveExp = archiveSheet.getRange(`B${firstArchiveRow}:F`).getValues(); //Previously archived expenses
  const archiveInc = archiveSheet.getRange(`I${firstArchiveRow}:M`).getValues(); //Previously archived income
  const transExp = transSheet.getRange(`B${firstTransRow}:F`).getValues(); //Current list of expenses
  const transInc = transSheet.getRange(`I${firstTransRow}:M`).getValues(); //Current list of income

  const nextRow = rng => rng.reduce((lastIndex,row,rowIndex) => { //Find the next empty row in a given range
    if (row.join('') !== '') return rowIndex + 1;
    return lastIndex;
  },0);

  const nextExpRow = nextRow(archiveExp) + firstArchiveRow; //Returns the next empty expenses row on the archive sheet
  const nextIncRow = nextRow(archiveInc) + firstArchiveRow; //Returns the next empty income row on the archive sheet
  const clearedExp = []; //Array to save the indecies of the expense transactions moved to archive
  const clearedInc = []; //Array to save the indecies of the income transactions moved to archive

  const monthExp = transExp.filter((v,i) => { //Filter the list of expenses to only keep the ones from the selected month
    if (!v[0]) return false;
    const date = new Date(v[0]).getTime();
    if (date >= selectedMonth && date <= endOfMonth) { //Check that they happened in the selected month
      if (date < new Date(new Date().toDateString()).getTime()) { //Only move transactions from before today
        if (v[2] !== '') { //Don't archive unreconciled transactions
          console.log('reconciled expense')
          clearedExp.push(i);
          return true;
        }
        else return false;
      }
      else return false;
    }
    else return false;
  });
  const monthInc = transInc.filter((v,i) => { //Filter the list of income to only keep the ones from the selected month
    if (!v[0]) return false;
    const date = new Date(v[0]).getTime();
    if (date >= selectedMonth && date <= endOfMonth) { //Check that they happened in the selected month
      if (date < new Date(new Date().toDateString()).getTime()) { //Only move transactions from before today
        if (v[2] !== '') { //Don't archive unreconciled transactions
          clearedInc.push(i);
          return true;
        }
        else return false;
      }
      else return false;
    }
    else return false;
  });

  archiveSheet.getRange(nextExpRow,2,monthExp.length,monthExp[0].length).setValues(monthExp); //Write the matching expenses to the archive
  archiveSheet.getRange(nextIncRow,9,monthInc.length,monthInc[0].length).setValues(monthInc); //Write the matching income to the archive

  archiveSheet.getRange(`B${firstArchiveRow}:F`).sort({column: 2, ascending: false}); //Sort the archived expenses list
  archiveSheet.getRange(`I${firstArchiveRow}:M`).sort({column: 9, ascending: false}); //Sort the archived income list

  clearedExp.forEach(v => transSheet.getRange(v + firstTransRow,2,1,5).clearContent()); //Clear the moved expenses from Transactions
  clearedInc.forEach(v => transSheet.getRange(v + firstTransRow,9,1,5).clearContent()); //Clear the moved income from Transactions

  transSheet.getRange(`B${firstTransRow}:F`).sort({column: 2, ascending: false}); //Sort the remaining expenses on Transactions
  transSheet.getRange(`I${firstTransRow}:M`).sort({column: 9, ascending: false}); //Sort the remaining income on Transactions

  console.log('Archived '+monthExp.length+' expense transactions');
  console.log('Archived '+monthInc.length+' income transactions');

  ss.getSheetByName(Utilities.formatDate(new Date(selectedMonth),Session.getScriptTimeZone(),"MMM yyyy")).hideSheet(); //Hide the archived month sheet, if it's shown
}

function advanceTransactions (param = 'To tomorrow') { //Use this function to reschedule any transactions that haven't yet been reconciled but are dated for today or earlier. The transactions will be moved to tomorrow and then the sheet will be sorted.
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Transactions');
  const values = sheet.getDataRange().getValues();

  const headerRow = values.findIndex(v => v.includes('Plan Amount')); //Find the index of the header row
  
  const dateCols = [values[headerRow].indexOf('Date'),values[headerRow].lastIndexOf('Date')]; //Index of both Date columns
  const actualCols = [values[headerRow].indexOf('Actual Amount'),values[headerRow].lastIndexOf('Actual Amount')]; //Index of both Actual Amount columns
  const today = new Date(new Date().toDateString()).getTime(); //Roundabout way to get the same date value as a date entered by text
  let needsSort = false;
  
  values.forEach((v,i) => {
    dateCols.forEach((d,j) => {
      if (v[actualCols[j]] === "") {
        if (param === 'To today') {
          if (new Date(v[d]).getTime() < today) {
            sheet.getRange(i+1,d+1).setValue(new Date(today));
            needsSort = true;
          }
        }
        else {
          if (new Date(v[d]).getTime() <= today) {
            sheet.getRange(i+1,d+1).setValue(new Date(today + (24 * 60 * 60 * 1000)));
            needsSort = true;
          }
        }
      }
    });
  });

  if (needsSort) sortTrans(); //Only sorts when transactions have been advanced
  //sortTrans(); //Sorts every time this function is ran
}

function sortTrans(sheetName = 'Transactions') {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  const expenseRange = sheet.getRange('B5:F');
  const incomeRange = sheet.getRange('I5:M');

  expenseRange.sort({column: 2, ascending: false});
  incomeRange.sort({column: 9, ascending: false});
}

function getHolidays(year = 2026) {
  const url = "https://date.nager.at/api/v4/PublicHolidays/" + year + "/US";
  
  try {
    const response = UrlFetchApp.fetch(url);
    const holidays = JSON.parse(response.getContentText());

    const holidayDates = holidays.map(holiday => {
      const dateParts = holiday.observedDate.split('-');
      if  (holiday.holidayTypes.includes('Public') && 
           (holiday.nationalHoliday === true || holiday.localName === 'Columbus Day')
          ) return new Date(dateParts[1] + "/" + dateParts[2] + "/" + dateParts[0]).toLocaleDateString();
      else return false;
    }).filter(Boolean);

    return holidayDates;
  }
  catch (e) {
    console.log("Error: " + e.toString());
    return [];
  }
}