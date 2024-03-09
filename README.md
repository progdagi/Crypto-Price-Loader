# Crypto-Price-Loader

## 1. Open crypto price Tracker Sheet Template and copy to your drive: 
https://docs.google.com/spreadsheets/d/1A1zFst2fkrpY_Ua5SYnKbW7xRIsidjq3TbS6rIrIfbw/edit?usp=sharing

## 2.price Tracker Sheet Template Column Design Requirements:
- Each column must include a cryptocurrency currency symbol, with the first cryptocurrency symbol starting from the "A2" cell.
- The "C2" cell should be designated as the initial cell for the crypto price loaded by Scrapit.
=> (must have those 3columns need in the sheet linke the template the above one )

## 3. From the menu Clikc on "Extensions -> App Scrapit"
- this will open new page of google app scrapit code editor wait until it load then ->
- remove all the codes in the page "make it empty "

## 4. Copy the below code and past to "google app scrapit code editor"
```js
/**
* Fetches the list of cryptocurrencies from the CoinMarketCap API.
* @returns {Object[]} An array containing information about cryptocurrencies.
* @throws {Error} Throws an error if the API request fails.
*/
function getCryptoCurrencyList() {
  try {
    let apiUrl = "https://api.coinmarketcap.com/data-api/v3/cryptocurrency/listing?limit=10000&sortBy=market_cap&sortType=desc&convert=USD&cryptoType=all&tagType=all&audited=false"
    let response = UrlFetchApp.fetch(apiUrl);
    let responseData = JSON.parse(response.getContentText());
    //"symbol": "BTC",
    //quotes->price
    return responseData["data"]["cryptoCurrencyList"];
  } catch (error) {
    Logger.log("Error fetching cryptocurrency list: " + error);
    throw new Error("Failed to fetch cryptocurrency list. Please try again later.");
  }
}
/**
* Gets the active sheet from the active spreadsheet.
* @returns {Sheet} The active sheet.
*/
function getActiveSheet() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  // Get the active sheet
  var sheet = spreadsheet.getActiveSheet()
  return sheet
}
/**
* Retrieves the symbols of cryptocurrencies from the specified sheet.
* @param {Sheet} sheet - The sheet containing the cryptocurrency symbols.
* @returns {string[]} An array containing cryptocurrency symbols.
*/
function getCryptoSymbols(sheet) {
  // Get the last row with content in the active sheet
  var lastRow = sheet.getLastRow();
  // Get the values of the first column (column A) for all rows
  var firstColumnRange = sheet.getRange("A2:A" + lastRow);
  var symbols = firstColumnRange.getValues();
  return symbols.flat(); //flat for changing multi-di array to single array
}
/**
* Appends the cryptocurrency price to the specified row in the sheet.
* @param {number} coinPrice - The price of the cryptocurrency.
* @param {Sheet} sheet - The sheet to append the price to.
* @param {number} row - The row number to append the price to.
*/
function appendCryptoPrice(coinPrice, sheet, row) {
  let price;
  if (coinPrice > 1) {
    price = coinPrice.toFixed(2)
  } else {
    price = coinPrice.toPrecision(4)
  }
  let cell = sheet.getRange("C" + row)
  cell.setValue(price)
}
/**
* Main function to load cryptocurrency prices into the active sheet.
*/
function priceLoader() {
  try {
    const sheet = getActiveSheet();
    const cryptoCurrencyList = getCryptoCurrencyList();
    const symbolsOnSheet = getCryptoSymbols(sheet)
    const filteredSymbol = [];
    const filteredCryptoCurrency = cryptoCurrencyList.filter(crypto => {
      if (symbolsOnSheet.includes(crypto.symbol) && !filteredSymbol.includes(crypto.symbol)) {
        filteredSymbol.push(crypto.symbol);
        return true;
      }
    });
    for (let crypto of filteredCryptoCurrency) {
      let coinPrice = crypto["quotes"][0]["price"];
      let row = symbolsOnSheet.indexOf(crypto.symbol) + 2;
      Logger.log(`Symbol: ${crypto.symbol}, Expected Price: ${coinPrice}`);
      // append it
      appendCryptoPrice(coinPrice, sheet, row);
    }
  } catch (error) {
    Logger.log("Error in main function: " + error);
    // Notify the user about the error
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
  }
}
/**
* Adds a custom menu to the Google Sheets UI.
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Crypto Price Loader')
    .addItem('Load Price', 'priceLoader')
    .addToUi();
}
```

## 5. save the file using ctr+s or cmd+s then close the page and back to the sheet and reload it
- you will see new custom menu: "Crypto Price Loader" then click it and select "Load Price" menu item any time you want a price update to append to price column

# how to add additional cryptocurrency to the sheet?
- add the crypto symbol and name in last row then follow the step: 5

Nb: you can add other columns but dont touch those 3 columns (id / symbol,name,price) columns 
  
