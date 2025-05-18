function fetchAndProcessTrades() {
    const token = '#########';  // Replace with your actual token
    const queries = {
        activity: '#######',  // Query ID for activity data
        tradeConfirmation: '#######'  // Query ID for trade confirmation data
    };

    // Fetch both activity and trade confirmation data
    Logger.log("Fetching activity data...");
    const activityData = fetchAndLogData(token, queries.activity, false);
    Logger.log(`Fetched ${activityData.length} activity trades.`);

    // Pause to avoid overloading the server
    Utilities.sleep(2000);

    Logger.log("Fetching trade confirmation data...");
    const tradeConfirmationData = fetchAndLogData(token, queries.tradeConfirmation, true);
    Logger.log(`Fetched ${tradeConfirmationData.length} trade confirmations.`);

    // Combine both sets of data
    const allTradesData = [...activityData, ...tradeConfirmationData];

    // Add combined trades data to the sheet
    Logger.log("Adding combined trades to the sheet...");
    addNewTrades(allTradesData);
}

// Function to fetch data from the API and parse XML
function fetchAndLogData(token, queryId, isTradeConfirmation = false) {
    const requestUrl = `https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService/SendRequest?t=${token}&q=${queryId}&v=3`;
    
    const response = UrlFetchApp.fetch(requestUrl, { muteHttpExceptions: true });
    const xmlContent = response.getContentText();
    
    return parseFlexResponse(xmlContent, token, isTradeConfirmation);
}

// Function to parse XML response and handle reference code logic
function parseFlexResponse(xml, token, isTradeConfirmation) {
    const document = XmlService.parse(xml);
    const root = document.getRootElement();

    if (root.getChild('Status')?.getText() === 'Fail') {
        const errorMessage = root.getChild('ErrorMessage')?.getText() || 'Unknown error';
        Logger.log("Request failed: " + errorMessage);
        return [];
    }

    const referenceCode = root.getChild('ReferenceCode')?.getText();
    if (referenceCode) {
        Logger.log("Reference Code: " + referenceCode);
        return fetchDataWithReference(token, referenceCode, isTradeConfirmation);
    } else {
        Logger.log("No reference code found in the response.");
    }

    return [];
}

// Function to fetch data using reference code and process it
function fetchDataWithReference(token, referenceCode, isTradeConfirmation) {
    const dataUrl = `https://gdcdyn.interactivebrokers.com/Universal/servlet/FlexStatementService.GetStatement?t=${token}&q=${referenceCode}&v=3`;
    
    const dataResponse = UrlFetchApp.fetch(dataUrl, { muteHttpExceptions: true });
    const dataXml = dataResponse.getContentText();

    return parseDataXML(dataXml, isTradeConfirmation);
}

// Function to parse the trade confirmation data from XML
function parseDataXML(xml, isTradeConfirmation) {
    const document = XmlService.parse(xml);
    const root = document.getRootElement();

    let tradesNode;
    if (isTradeConfirmation) {
        tradesNode = root.getChild('FlexStatements')?.getChild('FlexStatement')?.getChild('TradeConfirms');
        Logger.log('Looking for TradeConfirms node in trade confirmation data...');
    } else {
        tradesNode = root.getChild('FlexStatements')?.getChild('FlexStatement')?.getChild('Trades');
        Logger.log('Looking for Trades node in activity data...');
    }

    if (!tradesNode) {
        Logger.log('No trades node found in the response.');
        return [];
    }

    // Conditionally search for the correct trade elements based on the data type
    const trades = isTradeConfirmation 
        ? tradesNode.getChildren('TradeConfirm')  // For trade confirmation data
        : tradesNode.getChildren('Trade');        // For activity data

    Logger.log(`Found ${trades.length} trades in the response.`);

    return trades.map(trade => {
        const symbol = trade.getAttribute('symbol')?.getValue();

        // Skip "USD.SEK" trades as per your requirement
        if (symbol === "USD.SEK") {
            return null;
        }

        const quantity = parseInt(trade.getAttribute('quantity')?.getValue(), 10);

        // Handle price and tradePrice difference
        const tradePrice = isTradeConfirmation 
            ? parseFloat(trade.getAttribute('price')?.getValue())    // In trade confirmation data, it's "price"
            : parseFloat(trade.getAttribute('tradePrice')?.getValue());  // In activity data, it's "tradePrice"

        const dateTime = trade.getAttribute('dateTime')?.getValue();

        // Handle code for trade confirmation data, otherwise use openCloseIndicator
        const openCloseIndicator = isTradeConfirmation 
            ? trade.getAttribute('code')?.getValue()?.split(';')[0]  // In trade confirmation, it's "code" and we take the first part
            : trade.getAttribute('openCloseIndicator')?.getValue();  // In activity data, it's "openCloseIndicator"

        let tradeDate = 'N/A';
        let tradeTime = 'N/A';

        // Handle the order time to split it into date and time parts
        if (dateTime !== 'N/A') {
            const [datePart, timePart] = dateTime.split(';');
            tradeDate = `${datePart.slice(0, 4)}-${datePart.slice(4, 6)}-${datePart.slice(6, 8)}`;
            tradeTime = `${timePart.slice(0, 2)}:${timePart.slice(2, 4)}:${timePart.slice(4, 6)}`;
        }

        // Extract assetCategory and buySell (both are common in both XML types)
        const assetCategory = trade.getAttribute('assetCategory')?.getValue();
        const buySell = trade.getAttribute('buySell')?.getValue();

        Logger.log(`Trade found - Symbol: ${symbol}, Quantity: ${quantity}, Price: ${tradePrice}, Open/Close Status: ${openCloseIndicator}, AssetCategory: ${assetCategory}, Buy/Sell: ${buySell}, Date/Time: ${tradeDate} ${tradeTime}`);

        return {
            symbol,
            quantity,
            price: tradePrice,
            tradeDate,
            tradeTime,
            openCloseIndicator,  // This will now have the first letter of "code" or the full "openCloseIndicator"
            assetCategory,        // Add assetCategory
            buySell               // Add buySell
        };
    }).filter(trade => trade !== null);  // Filter out null trades (i.e., excluded trades)
}





// Function to add new trades to the sheet
function addNewTrades(tradesData) {
    const portfolioSheetName = 'Portfolio Data'; // Portfolio sheet name
    const portfolioSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(portfolioSheetName);

    if (!portfolioSheet) {
        Logger.log(`Sheet "${portfolioSheetName}" not found.`);
        return;
    }

    // Get the latest trade date and time from cell B1 in the Portfolio sheet
    const latestDateTimeCell = portfolioSheet.getRange('B1').getValue();  // Combined DateTime in B1
    Logger.log(`Latest DateTime from B1: ${latestDateTimeCell}`);

    let latestDateTime;

    // If the cell is empty, add all trades since there's no existing data
    if (!latestDateTimeCell) {
        Logger.log("No existing trade date and time found, adding all trades.");
        const sortedTrades = sortTradesByDateAndTime(tradesData);
        addTradesToSheet(portfolioSheet, sortedTrades);

        // Update B1 to the date and time of the latest trade (newest)
        const mostRecentTrade = sortedTrades[0];  // Get the newest trade now that it's sorted
        portfolioSheet.getRange('B1').setValue(`${mostRecentTrade.tradeDate}T${mostRecentTrade.tradeTime}`); // Update B1
        Logger.log(`Updated B1 with: ${mostRecentTrade.tradeDate}T${mostRecentTrade.tradeTime}`);
        return;
    }

    // Create Date object from the value in B1
    latestDateTime = new Date(latestDateTimeCell);
    
    if (isNaN(latestDateTime)) {
        Logger.log("Error: Latest DateTime is invalid.");
        return; // Exit if invalid date
    }

    // Log the latest date time we're comparing against
    Logger.log(`Comparing new trades with latest date/time: ${latestDateTime}`);

    // Filter the new trades based on whether they are newer than the latest recorded trade
    const newTrades = tradesData.filter(trade => {
        const tradeDateTimeStr = `${trade.tradeDate}T${trade.tradeTime}`;
        const tradeDateTime = new Date(tradeDateTimeStr);


        // Check if the trade date is greater than the latest date time
        const isNewerTrade = tradeDateTime > latestDateTime;

        return isNewerTrade;  
    });

    // Log how many new trades were found
    Logger.log(`Number of new trades found: ${newTrades.length}`);

    // Sort new trades from newest to oldest before adding
    const sortedNewTrades = sortTradesByDateAndTime(newTrades);

    // Add sorted new trades to the sheet
    if (sortedNewTrades.length > 0) {
        Logger.log(`Adding ${sortedNewTrades.length} new trades to the sheet.`);
        addTradesToSheet(portfolioSheet, sortedNewTrades);

        // Ensure the latest trade is updated after new trades have been added
        const mostRecentTrade = sortedNewTrades[0];  // Get the newest trade
        portfolioSheet.getRange('B1').setValue(`${mostRecentTrade.tradeDate}T${mostRecentTrade.tradeTime}`);  // Update B1
        Logger.log(`Updated B1 with: ${mostRecentTrade.tradeDate}T${mostRecentTrade.tradeTime}`);
    } else {
        Logger.log("No new trades to add.");
    }
}

// Function to sort trades by tradeDate and tradeTime in descending order (newest first)
function sortTradesByDateAndTime(trades) {
    return trades.sort((a, b) => {
        const dateTimeA = new Date(`${a.tradeDate}T${a.tradeTime}`);
        const dateTimeB = new Date(`${b.tradeDate}T${b.tradeTime}`);
        return dateTimeB - dateTimeA; // Sort descending, newest first
    });
}

// Helper function to add trades to the top of the sheet
function addTradesToSheet(sheet, trades) {
    const headerRow = 1;  // Assuming there's a header row
    const startRow = headerRow + 1;  // Start adding data after the header
    const startColumn = 3;  // Start adding data from column C

    // Map the trade data to array format for easy insertion into the sheet
    // Re-arrange the order to match: Symbol, Quantity, Price, Open/Close, Asset, Buy/Sell, Date, Time
    const tradeDataArray = trades.map(trade => [
        trade.symbol, 
        trade.quantity, 
        trade.price, 
        trade.openCloseIndicator,  // Move Open/Close before assetCategory
        trade.assetCategory,       // Asset column
        trade.buySell,             // Buy/Sell column
        trade.tradeDate,           // Move Date after Buy/Sell
        trade.tradeTime            // Move Time after Date
    ]);

    // Insert blank rows at the top of the sheet to make space for new trades
    sheet.insertRows(startRow, tradeDataArray.length);  // Insert as many rows as needed

    // Write new trade data into the newly inserted rows at the top
    sheet.getRange(startRow, startColumn, tradeDataArray.length, tradeDataArray[0].length)
        .setValues(tradeDataArray);
    
    Logger.log(`Added ${tradeDataArray.length} trades to the top of the sheet with rearranged columns.`);
}


