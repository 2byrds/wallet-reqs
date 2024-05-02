function populateDeps(sheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const depSheet = createOrClearSheet(ss, depSheetName);
  
  const nonDependencySheets = sheets.filter(sheet => sheet.getName() !== depSheetName);

  Logger.log(`Setting up row and column headers in the ${depSheetName} sheet.`);
  setupRowAndColumnHeaders(depSheet, nonDependencySheets);
  Logger.log(`Marking dependencies in the ${depSheetName} sheet.`);
  markDependencies(ss, depSheet, depSheet);
}

function setupRowAndColumnHeaders(depSheet, sheets) {
  const headers = sheets.map(sheet => [sheet.getName()]);
  depSheet.getRange(1, 2, 1, headers.length).setValues([headers.flat()]);
  depSheet.getRange(2, 1, headers.length, 1).setValues(headers);

  Logger.log(`${depSheet.getName()}: populating deps - sheet row and column headers have been set up.`);
}

// Assume markDependencies remains unchanged but follow similar logging enhancements
function markDependencies(ss, sheets, depSheet) {
  Logger.log("Starting to mark dependencies across sheets...");

  const sheetNames = sheets.map(sheet => sheet.getName());
  
  sheets.forEach(sheet => {
    let dependencyCount = 0; // Initialize a counter for dependencies
    
    // Log the beginning of processing for each sheet
    Logger.log(`${sheet.getName()}: Processing dependencies...`);
    
    const maxColumns = sheet.getMaxColumns();
    if (maxColumns < 2) {
      Logger.log(`${sheet.getName()}: skipped, has no headers.`);
      return; // Skip sheets with no possible headers
    }
    
    const formulas = sheet.getRange(1, 1, sheet.getMaxRows(), maxColumns).getFormulas();
    const formulaString = formulas.join();
    
    sheets.forEach((targetSheet, index) => {
      let dependencyFound = false;
      
      // Check for formula-based and header-based dependencies
      if (formulaString.includes(`'${targetSheet.getName()}'!`) || 
          formulaString.includes(`${targetSheet.getName()}:`)) {
        dependencyFound = true;
        dependencyCount++;
      }
      
      if (dependencyFound) {
        const targetRow = sheetNames.indexOf(sheet.getName()) + 2; // Adjust for headers
        const targetCol = index + 2; // Adjust for column B start
        depSheet.getRange(targetRow, targetCol).setValue('X').setBackground('green');
        // Log each dependency marked
        Logger.log(`${sheet.getName()}: depends on '${targetSheet.getName()}'`);
      }
    });
    
    // Log the result of dependency checking for each sheet
    if (dependencyCount > 0) {
      Logger.log(`${sheet.getName()}: Marked ${dependencyCount} dependencies.`);
    } else {
      Logger.log(`${sheet.getName()}: No dependencies found.`);
    }
  });

  Logger.log(`${sheet.getName()}: Finished marking dependencies.`);
}
