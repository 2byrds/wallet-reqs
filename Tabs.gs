const depMarker = ": ";

const acdc = 'Spec: ACDC';
const cesr = 'Spec: CESR';
const depSheetName = 'Dependency';
const identityProvider = 'App: vLEI: Identity provider'
const signReport = 'App: vLEI: Sign report'
const keri = 'Spec: KERI';
const keria = 'Impl: keria';
const keripy = 'Impl: KERIpy';
const signify = 'Protocol: Signify';
const signifyClient = 'App: Signify Client'
const signifypy = 'Impl: signifypy';
const signifyts = 'Impl: signify-ts';
const wallet = 'App: Wallet';

const knownSheets = new Map();
// const order = [cesr, keri, acdc, keripy, keria]
const order = [cesr, keri, acdc]
knownSheets.set(cesr,[])
knownSheets.set(keri,[])
knownSheets.set(acdc,[cesr,keri])
// knownSheets.set(keripy,[acdc, keri]);
// knownSheets.set(keria,[keripy,signify]);
// knownSheets.set(signifypy,[signify])
// knownSheets.set(signifyts,[signify])
// , , ,
//   , 'Architecture: KERIA', 'Dependency', 'Platform: Web',
//   'Platform: Mobile', 'Ecosystem: vLEI', : ,
//   'Service: Backer', 'Library: LMDB', 'Impl: KERIRust?', ,
//   , 'Language: Python', 'Protocol: KERI messages (KAPI?)',
//   , , 'Spec: did:webs',
//   'Spec: KERI', 'Use: vLEI: EBA', 'Spec: XBRL'
// };

function tabsMain() {
  Logger.log('Processing tabs.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const foundSheets = findAllSheets(ss);
  if(foundSheets.length != order.length) {
    const foundSheetNames = foundSheets.map(sheet => sheet.getName());
    Logger.log(`Found sheets doesn't match with the order array: \nfound: ${foundSheetNames.join(", ")} \nexpected order:${order.join(", ")}`)
  } 
  else {
    // Use the function to sort 'foundSheets'
    const sortedSheets = sortSheetsByOrder(foundSheets, order);
    const sortedSheetNames = sortedSheets.map(sheet => sheet.getName());
    // Log the sorted sheet names
    Logger.log(`Sorted Sheet Names: ${sortedSheetNames}`);

    sortedSheets.forEach(sheet => {
      trimEmptyCols(sheet)
      clearDepHeaders(sheet,depMarker); // Ensure dep headers are cleared
      trimEmptyCols(sheet)
      // populateSheets(ss,sheet);
    })
    // populateDeps(sortedSheets, depSheetName);
  }

}

function appendDepHeaders(sheet, depSheet, depHeaders) {
  depName = depSheet.getName()

  let emptyCell = getFirstEmptyHeader(sheet);
  Logger.log(`${sheet.getName()} w/ dependency ${depName}: starting to append depHeaders at column ${emptyCell.column}.`);

  if (depHeaders.length === 0) {
    Logger.log(`${sheet.getName()} w/ dependency ${depName}: skipping as it has no valid depHeaders to append.`);
    return;
  }

  const fontColorObj = depSheet.getRange(1, 2, 1, depHeaders.length).getFontColorObjects()[0];
  Logger.log(`${sheet.getName()} w/ dependency ${depName}: appending depHeaders.`);

  if (depHeaders.length > 0) {
    startCol = emptyCell.column
    graftHeaders(sheet, depName, depHeaders, startCol, fontColorObj[0]);

    // Correctly identify the range to be grouped
    // Ensure we only group the newly added headers, not including the group name's cell
    const groupStartCol = startCol; // Start grouping from the column after the group name
    const groupEndCol = startCol + headers.length; // End grouping at the last added header column
    // groupHeaders(sheet, groupStartCol, groupEndCol)
    // trimEmptyCols(sheet)
  }

  Logger.log(`${sheet.getName()} w/ dependency ${depName} completed setting up headers. Next free column: ${startCol + headers.length + 1}`);
}

function clearDepHeaders(sheet, depMarker) {
  Logger.log(`${sheet.getName()}: clearing dep headers from with marker '${depMarker}'.`);
  const lastHeaderCol = sheet.getLastColumn(); // Assuming lastCol is the actual last column that might contain data.
  
  if (lastHeaderCol >= 2) { // Ensure there are headers to check starting from column B.
    const headers = sheet.getRange(1, 2, 1, lastHeaderCol - 1).getValues()[0]; // Adjust to get the correct range.
    let rangesToDelete = [];
    let currentRangeStart = null;

    // Identify ranges of columns that include the depMarker.
    headers.forEach((header, i) => {
      if (header.toString().includes(depMarker)) {
        if (currentRangeStart === null) {
          currentRangeStart = i + 2; // 1-based index and starting from column B.
        }
      } else {
        if (currentRangeStart !== null) {
          rangesToDelete.push({ start: currentRangeStart, end: i + 2 }); // Correctly mark the end of the range.
          currentRangeStart = null;
        }
      }
    });

    // Check if the last column(s) with the marker should be included.
    if (currentRangeStart !== null) {
      lhcol = lastHeaderCol + 1
      Logger.log(`${sheet.getName()}: adding range to delete ${currentRangeStart} to ${lhcol+1}.`);
      rangesToDelete.push({ start: currentRangeStart, end: lhcol + 1 }); // Include the last column in the range if it's marked.
    }

    // Reverse the ranges to delete from last to first to avoid shifting issues.
    rangesToDelete.reverse().forEach(range => {
      const numColumnsToDelete = range.end - range.start;
      sheet.getRange(1,range.start,1,range.end).shiftColumnGroupDepth(-1) //remove any grouping before deleting
      if (numColumnsToDelete === 1) {
        Logger.log(`${sheet.getName()}: deleting column at position ${range.start}.`);
        sheet.deleteColumn(range.start);
      } else if (numColumnsToDelete > 1) {
        Logger.log(`${sheet.getName()}: deleting columns from ${range.start} to ${range.end - 1}, total ${numColumnsToDelete}.`);
        sheet.deleteColumns(range.start, numColumnsToDelete);
      }
    });
  } else {
    Logger.log(`${sheet.getName()}: No dep headers to clear from, or lastCol parameter is too low.`);
  }
}


function createOrClearSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`${sheet.getName()}: creating the sheet.`);
    sheet = ss.insertSheet(sheetName);
  } else {
    Logger.log(`${sheet.getName()}: clearing the sheet for fresh data.`);
    sheet.clear();
  }
  return sheet;
}

function findAllSheets(ss) {
  const sheets = ss.getSheets();
  Logger.log(`There are ${sheets.length} sheets`)

  const actualSheetNames = sheets.map(sheet => sheet.getName());
  let known = [...knownSheets.keys()];
  const missingSheets = known.filter(sheetName => !actualSheetNames.includes(sheetName));
  const extraSheets = actualSheetNames.filter(sheetName => !knownSheets.has(sheetName));
  
  Logger.log(missingSheets.length > 0 ? `Missing Sheets: '${missingSheets.join(", ")}'` : "No missing sheets identified.");
  Logger.log(extraSheets.length > 0 ? `Extra Sheets Detected: '${extraSheets.join(", ")}'` : "No extra sheets found.");
  
  foundSheets = sheets.filter(sheet => knownSheets.has(sheet.getName()))
  Logger.log(`Found ${foundSheets.length} valid sheets`)
  return foundSheets;
}

function getDepHeaders(sheet,depSheet) {
  depName = depSheet.getName()
  Logger.log(`${sheet.getName()} w/ dependency ${depName}: retrieving dependencies`);
  let sheetHeaders = sheet.getRange(1, 2, 1, sheet.getMaxColumns() - 1).getValues()[0];
  sheetHeaders = sheetHeaders.filter(header => header.length > 0);
  let depHeaders = depSheet.getRange(1, 2, 1, depSheet.getMaxColumns() - 1).getValues()[0];
  depHeaders = depHeaders.filter(header => header.length > 0);
  Logger.log(`${sheet.getName()} w/ dependency ${depName}: there are ${depHeaders.length} headers`);
  // Filter out headers that already include the depMarker to avoid duplication
  Logger.log(`${sheet.getName()} w/ dependency ${depName}: filtering dependency headers that already exist, current count ${depHeaders.length}.`);
  depHeaders = depHeaders.filter(dhead => sheetHeaders.indexOf(dhead) < 0);
  Logger.log(`${sheet.getName()} w/ dependency ${depName}: filtered dependency to ${depHeaders.length} headers: ${depHeaders.join(", ")}`);
  return depHeaders
}

function getFirstEmptyHeader(sheet,lastCol) {
  let firstEmptyIndex = 1
  let header = []
  if(lastCol > 2) {
    headers = sheet.getRange(1, 2, 1, lastCol).getValues()[0];
    firstEmptyIndex = headers.findIndex(header => !header) + 2; // Adjust for 1-based index and starting from column B
  }

  Logger.log(`${sheet.getName()}: first empty header cell at column ${firstEmptyIndex > 1 ? firstEmptyIndex : headers.length + 2}`);
  return { row: 1, column: firstEmptyIndex > 1 ? firstEmptyIndex : headers.length + 2 };
}

function groupHeaders(sheet, firstcol, lastcol) {
  // Only create a group if we added any headers
  if (lastcol > firstcol) {
    numCols = lastcol - firstcol
    Logger.log(`${sheet.getName()}: grouping headers from column ${firstcol} to ${lastcol}, ${numCols} total`);
    sheet.getRange(1, firstcol+1, 1, numCols).shiftColumnGroupDepth(1);
    const group = sheet.getColumnGroup(firstcol,1)
    if (group) {
      group.collapse();
      Logger.log(`${sheet.getName()}: collapsed group headers from column ${firstcol+1} to ${lastcol}.`);
    }
  }
}

function populateSheets(ss, sheet, lastCol) {
  Logger.log(`${sheet.getName()}: populating sheet`);
  deps = knownSheets.get(sheet.getName())
  Logger.log(`${sheet.getName()} w/ ${deps.length} dependencies ${deps.join(", ")}`);

  deps.forEach(dep => {
    const depSheet = ss.getSheetByName(dep);
    if (!depSheet) {
      Logger.log(`${sheet.getName()} w/ dependency ${depName}: skipping dependency, sheet does not exist.`);
      return;
    }

    depHeaders = getDepHeaders(sheet,depSheet);

    Logger.log(`${sheet.getName()}: adding extra column for group`)
    newCell = sheet.getRange(1,lastCol+1)
    newCell.setValue('')

    appendDepHeaders(sheet, depSheet, depHeaders, lastCol);
    return
  });
}

function graftHeaders(sheet, groupName, headers, startColumn, fontColor) {
  Logger.log(`${sheet.getName()} w/ dependency ${groupName}: beginning setup for '${groupName}' in starting at column ${startColumn}.`);

  // Prepare the header values and font colors for batch setting
  let headerValues = [[groupName]]; // Include group name as the first header
  let fontColors = [[fontColor]]; // Apply the same font color for all headers
  Logger.log(`${sheet.getName()} w/ dependency ${groupName}: With font ${fontColor}`);

  headers.forEach(header => {
    // Construct the full header name with the group name and depMarker
    const fullHeaderName = `${groupName}${depMarker}${header}`;
    headerValues[0].push(fullHeaderName);
    fontColors[0].push(fontColor);
  });

  // Apply the header values and font colors to the sheet in a batch
  const headersRange = sheet.getRange(1, startColumn, 1, headerValues[0].length);
  headersRange.setValues(headerValues);
  Logger.log(`${sheet.getName()} w/ dependency ${groupName}: headers - ${headerValues.join(",")}`);
  // Assuming your environment supports setting multiple font colors at once; otherwise, set them individually
  headersRange.setFontColors(fontColors);
  Logger.log(`${sheet.getName()} w/ dependency ${groupName}: fonts - ${fontColors.values().next().value}`);

  // Adjust column widths based on header content
  headers.forEach((header, index) => {
    const columnWidth = Math.max(100, header.length * 8);
    sheet.setColumnWidth(startColumn + 1 + index, columnWidth);
  });
  Logger.log(`${sheet.getName()} w/ dependency ${groupName}: Finished setting headers, fonts, and widths`);
}

function sortSheetsByOrder(foundSheets, order) {
  return foundSheets.sort((a, b) => {
    // Get the index of the sheet names in 'order'
    const indexA = order.indexOf(a.getName());
    const indexB = order.indexOf(b.getName());

    // Compare the indexes for sorting
    if (indexA > indexB) {
      return 1;
    }
    if (indexA < indexB) {
      return -1;
    }
    // If two sheets have the same name or are not found in 'order', they stay in the same order relative to each other
    return 0;
  });
}

function trimEmptyCols(sheet) {
  const maxColumns = sheet.getMaxColumns();
  const maxRows = sheet.getMaxRows();
  const data = sheet.getRange(1, 1, maxRows, maxColumns).getValues();

  let startColIndex = null;
  let emptyRanges = [];

  // Iterate through each column from last to first
  for (let col = maxColumns; col > 0; col--) {
    let isColumnEmpty = true;
    // Check each row in the current column
    for (let row = 0; row < maxRows; row++) {
      if (data[row][col - 1] !== '') { // Adjust for zero-based index
        isColumnEmpty = false;
        break;
      }
    }

    // If the column is empty and no start index is set, mark it
    if (isColumnEmpty && startColIndex === null) {
      startColIndex = col;
    }
    // If the column is not empty and a start index is set, we've found a range
    else if (!isColumnEmpty && startColIndex !== null) {
      emptyRanges.push({start: col + 1, end: startColIndex}); // Save the range
      startColIndex = null; // Reset for the next range
    }
  }

  // Check if the first column(s) are also empty
  if (startColIndex !== null) {
    emptyRanges.push({start: 1, end: startColIndex});
  }

  // Delete the empty ranges from last to first
  for (let i = emptyRanges.length - 1; i >= 0; i--) {
    let range = emptyRanges[i];
    let numColsToDelete = range.end - range.start + 1;
    Logger.log(`Deleting columns from ${range.start} to ${range.end} (${numColsToDelete} columns).`);
    sheet.deleteColumns(range.start, numColsToDelete);
  }

  Logger.log(`Deleted ${emptyRanges.length} ranges of empty columns.`);
}
