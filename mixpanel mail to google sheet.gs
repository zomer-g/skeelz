function extractStrings() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet 1");
    var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Output");

    if (!sheet) {
      throw new Error("Sheet 'Sheet 1' not found");
    }

    if (!outputSheet) {
      outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Output");
    }

    var data = sheet.getRange("B:B").getValues();
    var extractedData = [];

    for (var i = 0; i < data.length; i++) {
      var text = data[i][0];

      if (text && typeof text === "string") { // Check for non-empty string
        var startIndex = text.indexOf("Interested Candidates Report");
        if (startIndex !== -1) {
          var timestamp = extractTimestamp(text);
          var formattedTimestamp = formatDate(timestamp);

          var extractedStrings = [];
          startIndex = text.indexOf("/", startIndex) + 1;

          while (startIndex !== 0) {
            var endIndex = startIndex + 40;
            var extractedString = text.substr(startIndex, 40).trim();

            if (extractedStrings.length === 0) {
              extractedStrings.push(extractedString);
            } else if (extractedStrings.length === 1) {
              extractedStrings.push(extractedString);

              // Check if the combination of candidate_id and string2 already exists
              var combinationExists = checkCombinationExists(extractedData, extractedStrings[0], extractedStrings[1]);

              if (!combinationExists) {
                extractedData.push([extractedStrings[0], extractedStrings[1], formattedTimestamp]);
              }

              extractedStrings = [];
            }

            startIndex = text.indexOf("/", endIndex) + 1;
          }

          if (extractedStrings.length > 0) {
            var combinationExists = checkCombinationExists(extractedData, extractedStrings[0], "");

            if (!combinationExists) {
              extractedData.push([extractedStrings[0], "", formattedTimestamp]);
            }
          }
        }
      }
    }

    // Write data to Output sheet
    if (extractedData.length > 0) {
      var numRows = extractedData.length;
      var numColumns = extractedData[0].length;

      // Set column titles
      var columnTitles = ["candidate_id", "positionOfferId", "timestamp"];
      extractedData.unshift(columnTitles);

      outputSheet.clearContents(); // Clear existing contents in the sheet
      outputSheet.getRange(1, 1, numRows + 1, numColumns).setValues(extractedData);

      Logger.log("Data written to the Output sheet successfully.");
    } else {
      Logger.log("No data to write.");
    }
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
}

function extractTimestamp(text) {
  var startIndex = text.indexOf("Interested Candidates Report") + "Interested Candidates Report".length;
  var endIndex = startIndex + 24;
  var timestamp = text.substring(startIndex, endIndex).trim();
  return new Date(timestamp);
}

function formatDate(date) {
  var month = String(date.getMonth() + 1).padStart(2, "0");
  var day = String(date.getDate()).padStart(2, "0");
  var year = date.getFullYear();
  var hours = String(date.getHours()).padStart(2, "0");
  var minutes = String(date.getMinutes()).padStart(2, "0");

  return month + "/" + day + "/" + year + " " + hours + ":" + minutes;
}

function checkCombinationExists(data, candidate_id, positionOfferId) {
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === candidate_id && data[i][1] === positionOfferId) {
      return true;
    }
  }
  return false;
}
