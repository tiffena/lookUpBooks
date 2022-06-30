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
  // get current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // get Books sheet
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  var data = sheet.getDataRange().getValues();
  
  var lastRow = sheet.getLastRow();

  var isbn = data[lastRow-1][COLUMN_NUMBER_ISBN-1];
  var title = data[lastRow-1][COLUMN_NUMBER_TITLE-1];
  
  // We will only do a query if the ISBN is newly entered, i.e., the title field is empty. 
  // We will not query again if the user merely changes the information in the row such as 
  // edit author name, etc.
  if (isbn!='' && title == '' ) {
    // get book info from Google Books API with ISBN from last row  
    var bookInfo = lookupByISBN(isbn);

    if (bookInfo) {
      // extract book info
      data=extractBook(bookInfo);
      // update spreadsheet
      sheet=updateSheet(sheet, lastRow, data);
      // apply all pending spreadsheet changes
      SpreadsheetApp.flush();                 
    }
    else {
      // write error into Title cell
      sheet.getRange(lastRow, COLUMN_NUMBER_TITLE).setValue('Cannot locate book by this ISBN');
      return false;
    }
  }
}

// invoke Google Books API to get a JSON file of the book info
function lookupByISBN(isbn) {
  
  var url = API_URL + "&q=isbn:" + isbn;

  // Make a GET request using the query string constructed above
  var response = UrlFetchApp.fetch(url);
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

function extractBook(bookInfo) {
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

  var bookdata = {title:title, authors:authors, year:year, categories:categories, lang:lang, description:description};

  return bookdata;
}

function updateSheet(sheet,lastRow, data) {
  var s=sheet;        
  // write respective info into column cells of last row
  s.getRange(lastRow, COLUMN_NUMBER_TITLE).setValue(data['title']);
  s.getRange(lastRow, COLUMN_NUMBER_AUTHORS).setValue(data['authors']);
  s.getRange(lastRow, COLUMN_NUMBER_YEAR).setValue(data['year4']);
  s.getRange(lastRow, COLUMN_NUMBER_CATEGORIES).setValue(data['categories']);
  s.getRange(lastRow, COLUMN_NUMBER_LANG).setValue(data['lang']);
  s.getRange(lastRow, COLUMN_NUMBER_DESCRIPTION).setValue(data['description']);

  return s;               
}
