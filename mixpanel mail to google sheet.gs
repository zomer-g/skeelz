function createObjectsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet 1'); 
  const columnB = sheet.getRange('B:B').getValues(); 
  const TheTable = [['Cell', 'Date', 'Candidate_id', 'Position_id']]; // Add header row

  for (let i = 0; i < columnB.length; i++) {
    if (columnB[i][0]) { 
      const cellName = 'B' + (i+1).toString();
      
      const cellContent = columnB[i][0].toString(); 
      const datePattern = /[A-Z][a-z]{2} [A-Z][a-z]{2}\. \d{1,2}, \d{4}, \d{1,2}:\d{2} [A|P]M UTC/g;
      const dateMatches = cellContent.match(datePattern); 

      let dateString = '';
      if (dateMatches) {
        const dateObj = new Date(dateMatches[0]);
        dateString = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
      }

      const idPattern = /\b[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}\b/g; 
      const idMatches = cellContent.match(idPattern);
      
      if (idMatches) {
        for (let j = 0; j < idMatches.length; j += 2) {
          const pair = [
            cellName,
            dateString,
            idMatches[j],
            idMatches[j+1] ? idMatches[j+1] : ''
          ];
          TheTable.push(pair);
        }
      }
    }
  }

  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Output2'); 
  outputSheet.getRange(1, 1, TheTable.length, 4).setValues(TheTable);
}
