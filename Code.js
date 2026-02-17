/**
 * 
 * This script supports the Google Sheet that can be found here: https://docs.google.com/spreadsheets/d/11YmGiyljsHIa93hsYmDxWZ_L_ruKUmPMRlufNl3WJCA/copy
 * 
 * I use this project to track and forecast my finances. Check out README.md for more info!
 *  
 */

//Simple trigger that runs when the spreadsheet is first opened. I'm using it to help the user install the runOnEdit installable trigger that is necessary for this script to be triggered from the spreadsheet
function onOpen (e) {
  const props = PropertiesService.getDocumentProperties();
  const initialized = props.getProperty('init'); //We're using a script property to track if the trigger is installed
  if (initialized !== 'true') {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Sheet automation').addItem('Enable automation','addTriggers').addToUi();

    const lastNotify = JSON.parse(props.getProperty('lastNotify')) || 0;
    const oneDay = 24 * 60 * 60 * 1000; //Frequency in ms to suppress the alert to keep it from being too annoying
    const time = new Date().getTime();
    if (lastNotify < time - oneDay) { //Check if we should deliver another notification
      ui.alert(`This spreadsheet has built-in automation that makes several features work better. If you would like to enable all of the features, go to the 'Sheet automation' menu and select 'Enable automation'. Follow the Google prompts to authorize the advanced features.`);
      props.setProperty('lastNotify',JSON.stringify(time)); //Save the time of this notification
    }
  }
}

//This function will install the runOnEdit trigger and prevent the user from running too many times
function addTriggers () {
  ScriptApp.newTrigger('runOnEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create(); //Install the trigger
  const time = new Date().getTime();
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('init','true'); //Record that we've installed the trigger
  props.setProperty('simpleLog',JSON.stringify([time])); //Record the time in the edit logs to prevent the menu from erroneously appearing again
  props.setProperty('installableLog',JSON.stringify([time]));

  SpreadsheetApp.getUi().createMenu('Sheet automation').addItem('Done','onOpen').addToUi(); //Reset the menu
}

//Simple trigger to monitor whether or not the runOnEdit trigger is firing
function onEdit (e) {
  const time = new Date().getTime();
  const logLength = 5; //Number of entries to look back for matches, reduce to increase sensitivity, increase to be more sure about the determination
  const props = PropertiesService.getDocumentProperties();
  const log = JSON.parse(props.getProperty('simpleLog')) || []; //This is the log of times the simple trigger ran
  const log2 = JSON.parse(props.getProperty('installableLog')) || []; //This is the log of times the installable trigger ran
  const buffer = 2 * 60 * 1000; //Check to see if there's a matching entry within 2 minutes of each other
  let match = false;
  for (let i = 0; i < log.length; i++) {
    for (let j = 0; j < log2.length; j++) {
      if (Math.abs(log[i] - log2[j]) <= buffer) {
        match = true; //Record if we've found at least 1 timestamp on both logs that are within 2 minutes of each other
        break;
      }
    }
    if (match) break;
  }

  log.push(time); //Record this execution in the simple trigger log
  while (log.length > logLength) log.shift(); //Trim the log to the specified size
  props.setProperty('simpleLog',JSON.stringify(log)); //Save the log to the properties service

  if (!match) {
    props.deleteProperty('init'); //Delete the initialization flag if it appears that the runOnEdit function is not triggering
    onOpen(e); //Add the function to the menu and deliver an alert
  }
  console.log('match found?',match);
}

function runOnEdit(e) {
  const time = new Date().getTime();
  SpreadsheetApp.flush(); //Do a flush to make sure things are running smoothly before we start

  //Save a log of times this function is triggered so we can check that it's executing using a simple onEdit trigger
  const logLength = 5; //Number of execution times to save to the log
  const props = PropertiesService.getDocumentProperties();
  const log = JSON.parse(props.getProperty('installableLog')) || []; //Get previous log entries
  log.push(time); //Insert this execution into the log
  while (log.length > logLength) log.shift(); //Trim the log, if necessary
  props.setProperty('installableLog',JSON.stringify(log)); //Save the log to the properties service

  //const ss = SpreadsheetApp.getActive(); //Not used, commenting to save execution time
  const event = JSON.parse(JSON.stringify(e)); //Save the event variable in a way that's faster and easier to read
  console.log(event);

  const firstTransRow = 5; //This is the row of the first transaction on the Transactions tab

  /** @type {SpreadsheetApp.Sheet} */
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const actionSheets = ['Transactions','Archived Transactions','Recurring Transactions','Import Tool']; //List of sheets that have runOnEdit triggers

  if (!actionSheets.includes(sheetName)) return; //End execution if the end was not on a sheet in the list

  if (sheetName === 'Transactions') {
    const a1 = e.range.getA1Notation();
    if (a1 === 'K3') { //Advance late transactions
      if (!['To today','To tomorrow'].includes(e.value)) return; //Allow user to select between two options to advance the date

      e.range.setValue(''); //Reset the selector
      advanceTransactions(e.value);
    }
    if (a1 === 'L3') { //Sort transactions
      if (e.value !== 'TRUE') return;

      e.range.setValue(false); //Reset the checkbox
      sortTrans();
    }
    if (a1 === 'G3') { //Change selected to today
      if (e.value !== 'TRUE') return;

      e.range.setValue(false); //Reset the checkbox
      
      const today = new Date(new Date().toLocaleDateString()); //Set the date from a string so we can take time out of it to make it easier to compare dates
      const sheetVals = sheet.getDataRange().getValues(); //We need to load all of the values to search for checked boxes

      const firstIndex = sheetVals[firstTransRow - 2].indexOf('Date') - 1; //Get the first column labeled Date in the header row and decrease by 1
      const lastIndex = sheetVals[firstTransRow - 2].lastIndexOf('Date') - 1; //Get the last column labeled Date in the header row and decrease by 1
      sheetVals.forEach((v,i) => {
        if (v[firstIndex] === true || v[firstIndex] === 'TRUE') //Look for any rows in the Expenses column that's checked
          sheet.getRange(1 + i,1 + firstIndex,1,2).setValues([[false,today]]); //Reset the checkbox and change the date to today
      });
      sheetVals.forEach((v,i) => {
        if (v[lastIndex] === true || v[lastIndex] === 'TRUE') //Look for any rows in the Income column that's checked
          sheet.getRange(1 + i,1 + lastIndex,1,2).setValues([[false,today]]); //Reset the checkbox and change the date to today
      });
      
      //sortTrans();
    }
  }
  if (sheetName === 'Archived Transactions') {
    if (e.value !== 'TRUE') return;

    const a1 = e.range.getA1Notation();
    if (a1 === 'L3') {
      e.range.setValue(false); //Reset the checkbox
      sortTrans(sheetName);
    }
  }
  if (sheetName === 'Recurring Transactions') {
    if (e.value !== 'TRUE') return;

    const a1 = e.range.getA1Notation();
    if (a1 === 'K5') {
      e.range.setValue(false); //Reset the checkbox
      placeRecurring();
    }
    if (a1 === 'K10') {
      e.range.setValue(false); //Reset the checkbox
      archive();
    }
  }
  if (sheetName === 'Import Tool') {
    if (e.value !== 'TRUE') return;

    const a1 = e.range.getA1Notation();
    if (a1 === 'J3') {
      e.range.setValue(false); //Reset the checkbox
      importData();
    }
  }
}

//Place the list of planned recurring transactions onto the Transactions tab and also create a new sheet for the selected month.
function placeRecurring() {
  const firstTransRow = 5; //First row of transactions on the Transactions tab
  const firstRecurRow = 4; //First row of recurring transactions on the Recurring Transactions tab
  const ss = SpreadsheetApp.getActive();
  const recurSheet = ss.getSheetById(942392736) || ss.getSheetByName('Recurring Transactions');
  const transSheet = ss.getSheetById(1732160294) || ss.getSheetByName('Transactions');

  const month = recurSheet.getRange('L2').getValue();
  const year = recurSheet.getRange('L3').getValue();

  const props = PropertiesService.getDocumentProperties();
  const completedMonths = JSON.parse(props.getProperty('months')) || []; //Get list of previously processed months
  
  if (completedMonths.includes(month + year)) { //Check if this month has already been processed
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('The recurring transactions for this month have already been copied to the Transactions tab. If you continue, it may cause duplicate transactions to appear. Are you sure you want to continue?',ui.ButtonSet.YES_NO);
    if (response === ui.Button.NO) return console.log('user ended script execution');
  }

  const maxDates = { //Max valid dates for each month
    January: 31,
    February: year % 4 === 0 ? 29 : 28, //Account for leap year (won't work right in 2100, I don't care)
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

  const setDate = d => new Date(month + " " + d + ", " + year); //Get the date with no time
  const isWeekend = d => { //Determine if the date is on a weekend
    const weekday = Utilities.formatDate(d,Session.getScriptTimeZone(),"EEEE");
    return weekday === 'Saturday' || weekday === 'Sunday';
  };

  const holidays = getHolidays(year); //Gets a list of US holidays for the given year (update for other countries too?)
  const oneDay = 24 * 60 * 60 * 1000; //One day in miliseconds

  //Get the recurring Expenses and sort them by date
  const expenseData = recurSheet.getRange(`A${firstRecurRow}:D`).getValues();
  const expenses = !expenseData.join('') ? expenseData :
    expenseData.filter(r => r.join("") !== "").sort((a,b) => a[0] - b[0]).map(r => {
      r[0] = r[0] > maxDates[month] ? setDate(maxDates[month]) : setDate(r[0]); //Replace the day number with a full date being sure not to exceed the maximum number of days in the month
      while (holidays.includes(r[0].toLocaleDateString()) || isWeekend(r[0])) { //Check if date is on a weekend of holiday, repeat as necessary
        r[0] = new Date(r[0].getTime() - oneDay); //Move it earlier by one day if it's on a weekend or a holiday
      }
      r[1] = isNaN(r[1]) ? "" : r[1]; //Process the transaction without a fixed amount if no number is given for the amount
      r.splice(2,0,""); //Add another blank column to leave room for the actual amount on the Transactions tab
      return r;
    });
  //Repeat the same logic as above for the recurring Income list
  const incomeData = recurSheet.getRange(`F${firstRecurRow}:I`).getValues();
  const income = !incomeData.join('') ? incomeData :
    incomeData.filter(r => r.join("") !== "").sort((a,b) => a[0] - b[0]).map(r => {
      r[0] = r[0] > maxDates[month] ? setDate(maxDates[month]) : setDate(r[0]);
      while (holidays.includes(r[0].toLocaleDateString()) || isWeekend(r[0])) {
        r[0] = new Date(r[0].getTime() - oneDay);
      }
      r[1] = isNaN(r[1]) ? "" : r[1];
      r.splice(2,0,"");
      return r;
    });

  const transExp = transSheet.getRange(`B${firstTransRow}:F`).getValues(); //Get the list of expenses already on the Transactions tab
  const transInc = transSheet.getRange(`I${firstTransRow}:M`).getValues(); //Get the list of income already on the Transactions tab

  const nextRow = rng => rng.reduce((lastIndex,row,rowIndex) => { //Find the next open row in the list of transactions
    if (row.join('') !== '') return rowIndex + 1;
    return lastIndex;
  },0);

  const nextExpRow = nextRow(transExp) + firstTransRow; //Get the next empty row in the list of Expenses
  const nextIncRow = nextRow(transInc) + firstTransRow; //Get the next empty row in the list of Income

  if (!completedMonths.includes(month + year)) completedMonths.push(month + year); //Add this month to the list of processed months
  props.setProperty('months',JSON.stringify(completedMonths)); //Save the list of processed months to the properties service

  if (expenses.join(''))
    transSheet.getRange(nextExpRow,2,expenses.length,expenses[0].length).setValues(expenses); //Write the Expenses to the Transactions tab
  if (income.join(''))
    transSheet.getRange(nextIncRow,9,income.length,income[0].length).setValues(income); //Write the list of Income to the Transactions tab

  transSheet.getRange(`B${firstTransRow}:F`).sort({column: 2, ascending: false}); //Sort the list of Expenses
  transSheet.getRange(`I${firstTransRow}:M`).sort({column: 9, ascending: false}); //Sort the list of Income

  const prevSheetName = Utilities.formatDate(new Date(setDate(1).getTime() - (24 * 60 * 60 * 1000)),Session.getScriptTimeZone(),"MMM yyyy"); //Get the tab name of the month prior to the one that's being processed
  const nextSheetName = Utilities.formatDate(setDate(1),Session.getScriptTimeZone(),"MMM yyyy"); //Get the tab name of the month being processed

  let nextMonthSheet = ss.getSheetByName(nextSheetName); //Get the sheet with the name of the processed month, if it exists already
  if (!nextMonthSheet) { //Check if the processed month already has a sheet, continue if it doesn't exist already
    const prevMonthSheet = ss.getSheetByName(prevSheetName); //Get the prior month's sheet
    nextMonthSheet = prevMonthSheet.copyTo(ss).setName(nextSheetName); //Copy the prior month's sheet and rename it
    nextMonthSheet.getRange('C7').setValue(setDate(1)); //Set the start date of the new month's sheet
    nextMonthSheet.getRange('O12').setValue(""); //Reset the notes for the new sheet
    nextMonthSheet.getRange('L8').setValue(""); //Clear the prior month's starting value in case it was set manually
    nextMonthSheet.setActiveRange(nextMonthSheet.getRange('A1')); //Activate the new sheet
    ss.moveActiveSheet(2); //Sort the new sheet ahead of the previous
  }
}

//Move transactions for the selected month the the Archived Transactions tab to reduce clutter
function archive () {
  const firstTransRow = 5; //First row of the transactions on the Transactions tab
  const firstArchiveRow = 5; //First row of the transactions on the Archived Transactions tab
  const ss = SpreadsheetApp.getActive();
  const recurSheet = ss.getSheetById(942392736) || ss.getSheetByName('Recurring Transactions'); //Recurring transactions sheet
  const transSheet = ss.getSheetById(1732160294) || ss.getSheetByName('Transactions'); //Transactions sheet
  const archiveSheet = ss.getSheetById(1366470246) || ss.getSheetByName('Archived Transactions'); //Archived Transactions sheet

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
  const nextMonth = new Date(year, monthIndex[month] + 1,1).getTime(); //Return the 1st day of the month after the selected month in miliseconds
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

//Use this function to reschedule any transactions that haven't yet been reconciled but are dated for today or earlier. The transactions will be moved to tomorrow (or today) and then the sheet will be sorted.
function advanceTransactions (param = 'To tomorrow') {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Transactions');
  const values = sheet.getDataRange().getValues();

  const headerRow = values.findIndex(v => v.includes('Plan Amount')); //Find the index of the header row
  
  const dateCols = [values[headerRow].indexOf('Date'),values[headerRow].lastIndexOf('Date')]; //Index of both Date columns
  const actualCols = [values[headerRow].indexOf('Actual Amount'),values[headerRow].lastIndexOf('Actual Amount')]; //Index of both Actual Amount columns
  const today = new Date(new Date().toDateString()).getTime(); //Roundabout way to get the same date value as a date entered by text
  let needsSort = false;
  
  values.forEach((v,i) => {
    dateCols.forEach((d,j) => { //Runs for the list of Expenses and then the list of Income
      if (v[actualCols[j]] === "") { //Only run for unreconciled transactions
        if (( param === 'To today' && new Date(v[d]).getTime() < today ) ||
            ( param === 'To tomorrow' && new Date(v[d]).getTime() <= today)) { //Check if transaction is late
          sheet.getRange(i+1,d+1).setValue(new Date(today + (param === 'To today' ? 0 : (24 * 60 * 60 * 1000)))); //Advance late transactions
          needsSort = true; //Indicate if we need to sort
        }
      }
    });
  });

  if (needsSort) sortTrans(); //Only sorts when transactions have been advanced
  //sortTrans(); //Sorts every time this function is ran
}

//Process the data from the Import Tools tab to the Transactions tab
function importData () {
  const ss = SpreadsheetApp.getActive();
  const inSheet = ss.getSheetByName('Import Tool');
  const outSheet = ss.getSheetByName('Transactions');
  const inValues = inSheet.getDataRange().getValues();
  const outValues = outSheet.getDataRange().getValues();
  const headerRow = outValues.findIndex(v => v.includes('Plan Amount')); //Find the index of the header row
  const inHeaders = inValues.shift(); //Save the headers for the import data
  const outHeaders = outValues[headerRow]; //Save the headers for the transaction data

  const inExp = inValues.filter(v => { //Extract the expenses
    return v[inHeaders.indexOf('Withdrawal')];
  });
  const inInc = inValues.filter(v => { //Extract the income
    return v[inHeaders.indexOf('Deposit')];
  });

  const outExp = inExp.map(v => { //Format the expenses to match the Transactions tab
    return [
      v[inHeaders.indexOf('Date')],
      v[inHeaders.indexOf('Withdrawal')],
      v[inHeaders.indexOf('Withdrawal')],
      v[inHeaders.indexOf('Description')]
    ];
  });
  const outInc = inInc.map(v => { //Format the income to match the Transactions tab
    return [
      v[inHeaders.indexOf('Date')],
      v[inHeaders.indexOf('Deposit')],
      v[inHeaders.indexOf('Deposit')],
      v[inHeaders.indexOf('Description')]
    ];
  });

  const transData = [...outValues]; 
  transData.splice(0,headerRow + 1); //Get a new array with only the Transaction data
  const expCols = [ //Find the column indecies for the expenses
    outHeaders.indexOf('Date'),
    outHeaders.indexOf('Plan Amount'),
    outHeaders.indexOf('Actual Amount'),
    outHeaders.indexOf('Description')
  ];
  const incCols = [ //Find the column indecies for the income
    outHeaders.lastIndexOf('Date'),
    outHeaders.lastIndexOf('Plan Amount'),
    outHeaders.lastIndexOf('Actual Amount'),
    outHeaders.lastIndexOf('Description')
  ];

  const expTrans = selectColumns(transData, expCols); //Get arrays of just the expense data
  const incTrans = selectColumns(transData, incCols); //Get arrays of just the income data

 const nextRow = rng => rng.reduce((lastIndex,row,rowIndex) => { //Find the next empty row in a given range
    if (row.join('') !== '') return rowIndex + 1;
    return lastIndex;
  },0);

  const nextExpRow = nextRow(expTrans) + headerRow + 2; //Returns the next empty expenses row on the archive sheet
  const nextIncRow = nextRow(incTrans) + headerRow + 2; //Returns the next empty income row on the archive sheet

  outSheet.getRange(nextExpRow,outHeaders.indexOf('Date')+1,outExp.length,outExp[0].length).setValues(outExp);
  outSheet.getRange(nextIncRow,outHeaders.lastIndexOf('Date')+1,outInc.length,outInc[0].length).setValues(outInc);

  const inCols = inHeaders.filter(Boolean).length; //Get the number of columns in the import sheet
  inSheet.getRange(2,1,inSheet.getLastRow(),inCols).clear(); //Clears the imported data

  sortTrans(); //Sort the Transactions tab
}

//Sort the transactions by date
function sortTrans(sheetName = 'Transactions') {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  const expenseRange = sheet.getRange('B5:F');
  const incomeRange = sheet.getRange('I5:M');

  expenseRange.sort({column: 2, ascending: false});
  incomeRange.sort({column: 9, ascending: false});
}

//Gets a list of US holidays for the given year
function getHolidays(year = 2026) {
  const url = "https://date.nager.at/api/v4/PublicHolidays/" + year + "/US";
  
  try {
    const response = UrlFetchApp.fetch(url);
    const holidays = JSON.parse(response.getContentText());

    const holidayDates = holidays.map(holiday => {
      const dateParts = holiday.observedDate.split('-'); //Splits the date to [year, month, day]
      if  (holiday.holidayTypes.includes('Public') && 
           (holiday.nationalHoliday === true || holiday.localName === 'Columbus Day') //The latest version doesn't list Columbus Day as a national holiday so we'll manually check that one
          ) return new Date(dateParts[1] + "/" + dateParts[2] + "/" + dateParts[0]).toLocaleDateString(); //Get the date as a date object
      else return false;
    }).filter(Boolean); //Filter out non-national holidays

    return holidayDates;
  }
  catch (e) {
    console.log("Error: " + e.toString());
    return [];
  }
}

// Function to keep only columns at selected indices
function selectColumns(dataArray, columnIndicesToSelect) {
    return dataArray.map(row => {
        // Use filter to select elements whose index is in the list of indices given
        return row.filter((cell, cellIndex) => {
            return columnIndicesToSelect.includes(cellIndex);
        });
    });
}