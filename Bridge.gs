// 1. Build the menu when the workbook opens
function onOpen() {
  SWGOHInventoryParser.onOpen();
}

// 2. Bridge function for the custom menu button click
function showImportDialog() {
  SWGOHInventoryParser.showImportDialog();
}

// 3. Bridge function for the HTML Dialog data submission
function processSWGOHData(jsonString) {
  // Forwards the JSON payload to the library and returns the success/error message
  return SWGOHInventoryParser.processSWGOHData(jsonString); 
}
