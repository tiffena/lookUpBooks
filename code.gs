const API_URL = "https://www.googleapis.com/books/v1/volumes?country=US";

// Our Google Sheets sheet
const SHEET_NAME = 'Books';
// Our book info columns 
const COLUMN_NUMBER_ISBN = 2;
const COLUMN_NUMBER_TITLE = 3;
const COLUMN_NUMBER_AUTHORS = 4;
const COLUMN_NUMBER_YEAR = 5;
const COLUMN_NUMBER_CATEGORIES = 6;
const COLUMN_NUMBER_LANG = 7;
const COLUMN_NUMBER_DESCRIPTION = 8;

function main() {
    
  // log starting of the script
  Logger.log('Script has started');

  // get current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // get Books sheet
  var sheet = ss.getSheetByName(SHEET_NAME);

  // get last row in the sheet
  var lastRow = sheet.getLastRow();
     
  // get ISBN (column 2) of the last row 
  var data = sheet.getRange(lastRow, 2).getValues();
  var isbn = data[0][0];
  Logger.log('Looking Up ISBN:  ' + isbn);
  
  // get book info from Google Books API with lastrow's ISBN
  var bookInfo = lookupByISBN(isbn);
  // verify the call is successful
  if (bookInfo) {
        // extract book info
        var title = (bookInfo["volumeInfo"]["title"]);            
        var authors = (bookInfo["volumeInfo"]["authors"]);
        var year = (bookInfo["volumeInfo"]["publishedDate"]);
        // extract yyyy from published date
        if (year != null ) {
          var year4 = year.substring(0,4);} 
        var categories = (bookInfo["volumeInfo"]["categories"]);
        var lang = (bookInfo["volumeInfo"]["language"]);   
        var description = (bookInfo["volumeInfo"]["description"]);  
        // searchInfo.textSnippet is shorter than description
        // If this also exists, overwrite the description above
        if (bookInfo["searchInfo"] != null){
          if (bookInfo["searchInfo"]["textSnippet"] != null) {
            description = (bookInfo["searchInfo"]["textSnippet"]);  } }    
          
        // write respective info into column cells of last row
        sheet.getRange(lastRow, COLUMN_NUMBER_TITLE).setValue(title);
        sheet.getRange(lastRow, COLUMN_NUMBER_AUTHORS).setValue(authors);
        sheet.getRange(lastRow, COLUMN_NUMBER_YEAR).setValue(year4);
        sheet.getRange(lastRow, COLUMN_NUMBER_CATEGORIES).setValue(categories);
        sheet.getRange(lastRow, COLUMN_NUMBER_LANG).setValue(lang);
        sheet.getRange(lastRow, COLUMN_NUMBER_DESCRIPTION).setValue(description);
                
        // apply all pending spreadsheet changes
        SpreadsheetApp.flush();        
      } 
      else {
        // write error into Title cell and return false value
        sheet.getRange(lastRow, COLUMN_NUMBER_TITLE).setValue('Error finding ISBN data. See Logs');
        return false;
      }
    
  // Log completion of the script
  Logger.log('Script finished');
}

// invoke Google Books API to get a JSON file of the book info
function lookupByISBN(isbn) {
  
  var url = API_URL + "&q=isbn:" + isbn;

  // Make a GET request using the query string constructed above
  var response = UrlFetchApp.fetch(url);
  // 
  var results = JSON.parse(response);
  
  if (results.totalItems) {
    // for multiple results
    // get only the 1st item that should be closest match
    var book = results.items[0];
    Logger.log("Book found")
    return book;
  } 

  // our search returns no result
  return false;
}

 
