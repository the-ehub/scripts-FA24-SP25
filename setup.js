/**
 * Displays options under the "Activities" option in the navigation bar
 */
function onOpen() {
  let menu = SpreadsheetApp.getUi().createMenu('Activities')
      .addItem('Find Available Students', 'showSidebar')
      .addToUi();
}

/**
 * Shows the sidebar
 */
function showSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Find Available Students');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
