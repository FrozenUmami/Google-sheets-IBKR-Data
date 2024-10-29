function processTrades() {
  // Define the source sheet ("Portfolio Data"), result sheet ("Results"), and the new journal sheet ("Journal")
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Portfolio Data');
  var resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  var journalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Journal');  // New "Journal" sheet

  var dataRange = sourceSheet.getRange(2, 3, sourceSheet.getLastRow() - 1, 8);
  var data = dataRange.getValues();

  data.reverse();

  Logger.log("Data Length: " + data.length);  // Log the number of rows retrieved

  // Loop through each trade from "Portfolio Data"
  for (var i = 0; i < data.length; i++) {
    var rowData = data[i];
    var symbol = rowData[0];
    var quantity = rowData[1];
    var entryPrice = rowData[2];
    var position = rowData[3]; // "O" or "C"
    var status = "Open"; // Initial status for a new trade

    var openQuantity = 0;
    var closeQuantity = 0;
    var sumEntryPrice = 0;
    var sumExitPrice = 0;
    var qtyLeft = quantity;

    // Capture the latest trade's entry date and time
    var latestTradeDate = rowData[6]; // Assuming this is the date from Portfolio Data
    var latestTradeTime = rowData[7]; // Assuming this is the time from Portfolio Data

    // Check if the result sheet has any data beyond the header
    var lastResultRow = resultSheet.getLastRow();
    var resultData = [];

    if (lastResultRow > 1) {  // Only fetch result data if there are rows beyond the header
      resultData = resultSheet.getRange(2, 3, lastResultRow - 1, 12).getValues(); 
    }

    var foundOpenTrade = false;

    // Check for existing open trades with the same symbol
    for (var j = 0; j < resultData.length; j++) {
      var resultRow = resultData[j];
      var resultSymbol = resultRow[0];
      var resultStatus = resultRow[1];
      var resultQtyLeft = resultRow[2];
      var resultOpenQty = resultRow[3];
      var resultCloseQty = resultRow[4];
      var resultSumEntryPrice = resultRow[5];
      var resultSumExitPrice = resultRow[6];
      var resultEntryDate = new Date(resultRow[10]); // Entry Date (Column M)
      var resultEntryTime = new Date(resultRow[11]); // Entry Time (Column N)

      // Combine date and time for comparison
      var resultEntryDateTime = new Date(resultEntryDate.getFullYear(), resultEntryDate.getMonth(), resultEntryDate.getDate(),
                                         resultEntryTime.getHours(), resultEntryTime.getMinutes(), resultEntryTime.getSeconds());

      if (resultSymbol === symbol && resultStatus === "Open") {
        foundOpenTrade = true;

        // Update existing open trade
        if (position === "O") {
          // Add to open quantity and sum of entry prices
          resultRow[2] = resultQtyLeft + quantity;
          resultRow[3] = resultOpenQty + quantity;
          resultRow[5] = resultSumEntryPrice + (quantity * entryPrice);
        } else if (position === "C") {
          // Add to close quantity and sum of exit prices
          resultRow[2] = resultQtyLeft + quantity;
          resultRow[4] = resultCloseQty + quantity;
          resultRow[6] = resultSumExitPrice + (quantity * entryPrice);
        }

        // Update status if quantity left is 0
        if (resultRow[2] === 0) {
          resultRow[1] = "Closed";

          // Calculate Avg Entry Price and Avg Exit Price when status is "Closed"
          if (Math.abs(resultRow[3]) > 0) { // Open Quantity exists
            resultRow[7] = Math.abs(resultRow[5] / Math.abs(resultRow[3]));  // Avg Entry Price as positive
          }

          if (Math.abs(resultRow[4]) > 0) { // Close Quantity exists
            resultRow[8] = Math.abs(resultRow[6] / Math.abs(resultRow[4]));  // Avg Exit Price as positive
          }
        }

        // Add latest trade date & time to columns O (14th column) and P (15th column)
        resultRow[12] = latestTradeDate;  // Latest Trade Date in Column O
        resultRow[13] = latestTradeTime;  // Latest Trade Time in Column P


        // Write the updated row back to the "Results" sheet
        resultSheet.getRange(j + 2, 3, 1, 14).setValues([resultRow]); 

        // Also update the "Journal" sheet
        var journalRow = [
          resultEntryDate,    // Entry Date in column A
          resultEntryTime,    // Entry Time in column B
          resultRow[1],       // Status in column C
          resultSymbol,       // Ticker in column D
          resultOpenQty,      // Open Quantity in column E
          resultRow[7],       // Avg Entry in column G
          resultRow[8]        // Avg Exit in column M
        ];

        journalSheet.getRange(j + 11, 1, 1, 7).setValues([journalRow]);  // Start inserting from row 11 in the "Journal" sheet

        break;
      }
    }

    // If no open trade found and the position is "C", skip this trade (can't close a trade that isn't open)
    if (!foundOpenTrade && position === "C") {
      Logger.log("Skipping 'C' position for " + symbol + " as no open position exists.");
      continue;
    }

    // If no open trade found, check if the new trade is newer than the one at row 2
    if (!foundOpenTrade) {
      var entryTime = new Date(0, 0, 0, rowData[7].getHours(), rowData[7].getMinutes(), rowData[7].getSeconds());
      var newEntryDate = new Date(rowData[6]); // Current new trade's entry date
      var newEntryDateTime = new Date(newEntryDate.getFullYear(), newEntryDate.getMonth(), newEntryDate.getDate(),
                                       entryTime.getHours(), entryTime.getMinutes(), entryTime.getSeconds());

      Logger.log("New Trade Entry DateTime: " + newEntryDateTime);

      // Fetch the entry date and time of the trade in row 2 of "Results" sheet
      var topTradeEntryDateTime;
      if (lastResultRow > 1) {
        var topTradeEntryDate = resultSheet.getRange(2, 13).getValue(); // Get entry date in row 2
        var topTradeEntryTime = resultSheet.getRange(2, 14).getValue(); // Get entry time in row 2

        // Ensure that topTradeEntryDate and topTradeEntryTime are valid dates
        if (topTradeEntryDate instanceof Date && !isNaN(topTradeEntryDate) &&
            topTradeEntryTime instanceof Date && !isNaN(topTradeEntryTime)) {
          topTradeEntryDateTime = new Date(topTradeEntryDate.getFullYear(), topTradeEntryDate.getMonth(), topTradeEntryDate.getDate(),
                                            topTradeEntryTime.getHours(), topTradeEntryTime.getMinutes(), topTradeEntryTime.getSeconds());
        }
      }

      // Check if the top trade date is blank or missing
      if (!topTradeEntryDateTime || isNaN(topTradeEntryDateTime)) {
        Logger.log("Top trade date is missing or blank. Inserting new trade immediately.");
      } else {
        Logger.log("Top Trade Entry DateTime in Row 2: " + topTradeEntryDateTime);

        // Log the date and time comparison results
        if (newEntryDateTime <= topTradeEntryDateTime) {
          Logger.log("New trade is NOT newer. Skipping insertion.");
          continue; // Skip this trade as it's not newer
        }
      }

      // Insert new trade logic goes here (since no open trade was found or top trade date is missing)
      resultSheet.insertRowBefore(2);

      // Create the new trade row without "Position" and "Buy/Sell"
      var newRow = [
        symbol,           // Symbol (Column C)
        status,           // Status (Column D)
        qtyLeft,          // Qty Left (Column E)
        (position === "O" ? quantity : 0), // Open Quantity (Column F)
        (position === "C" ? quantity : 0), // Close Quantity (Column G)
        (position === "O" ? entryPrice * quantity : 0), // Sum Entry Price (Column H)
        (position === "C" ? entryPrice * quantity : 0), // Sum Exit Price (Column I)
        "",               // Avg Entry Price (Column J)
        "",               // Avg Exit Price (Column K)
        rowData[4],       // Asset Type (Column L)
        rowData[6],       // Entry Date (Column M)
        entryTime,         // Entry Time (Column N)
        latestTradeDate,  // Latest Trade Date (Column O)
        latestTradeTime   // Latest Trade Time (Column P)
      ];

      // Insert the new trade row into the "Results" sheet
      resultSheet.getRange(2, 3, 1, 14).setValues([newRow]);

      // Insert the new trade row into the "Journal" sheet
      var journalRow = [
        rowData[6],        // Entry Date in column A
        entryTime,         // Entry Time in column B
        status,            // Status in column C
        symbol,            // Ticker in column D
        (position === "O" ? quantity : 0), // Open Quantity in column E
        (position === "O" ? entryPrice : ""), // Avg Entry in column G
        ""                 // Avg Exit (not available for new entries)
      ];

      journalSheet.insertRowBefore(11);  // Insert before row 11
      journalSheet.getRange(11, 1, 1, 7).setValues([journalRow]);  // Insert starting from row 11 in the "Journal" sheet
    }
  }

  // Format the Entry Time column (now Column N) as "hh:mm:ss"
  if (lastResultRow > 1) {
    resultSheet.getRange(2, 12, resultSheet.getLastRow() - 1, 1).setNumberFormat("hh:mm:ss");  // Adjusted for new column (12 instead of 14)
  }

  Logger.log("Trade processing complete.");
}
