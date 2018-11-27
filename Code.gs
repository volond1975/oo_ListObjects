//http://www.javascript-spreadsheet-programming.com/2013/01/object-oriented-javascript-part-2.html#blog
//http://www.javascript-spreadsheet-programming.com/2015/04/using-google-spreadsheet-named-ranges.html#blog
//https://php-academy.kiev.ua/blog/required-javascript-reading #book
//https://www.planetaexcel.ru/techniques/2/136/#blog
//#book https://books.google.com.ua/books?id=K4mUCwAAQBAJ&pg=PA154&lpg=PA154&dq=listobject+google+spreadsheet&source=bl&ots=8ffgVU6jGn&sig=dJ2LvoS6X6RbUUnr_TrvGli0GGk&hl=ru&sa=X&ved=2ahUKEwiZ5dmIg_LeAhVotosKHZLLCjMQ6AEwBXoECAQQAQ#v=onepage&q=listobject%20google%20spreadsheet&f=false

'use strict';
/*global SpreadsheetApp: false */
DATATABLE="DataTable_"
function HeaderRow(spreadsheet, sheetName, headerRowNumber, startColumnNumber, columnTitles, overwritePrevious) {
  if (arguments.length !== 6) {
    throw {'name': 'Error',
           'message': '"HeaderRow()" constructor function requires 6 arguments!'};
  }
  this.spreadsheet = spreadsheet;
  this.sheetName = sheetName;
  this.headerRowNumber = headerRowNumber;
  this.startColumnNumber = startColumnNumber;
  this.columnTitles = columnTitles;
  this.overwritePrevious = overwritePrevious;
  this.sheet = this.spreadsheet.getSheetByName(this.sheetName);
  this.columnTitleCount = this.columnTitles.length;
  this.headerRowRange = this.sheet.getRange(this.headerRowNumber,
                                            this.startColumnNumber,
                                            1,
                                            this.columnTitleCount);
  this.headerRowRange.setFontWeight('normal');
  this.headerRowRange.setFontStyle('normal');
  this.addColumnTitlesToHeaderRow();
}

HeaderRow.prototype = {
  constructor: HeaderRow,
  freezeHeaderRow: function () {
    var sheet = this.sheet;
    sheet.setFrozenRows(this.headerRowNumber);
  },
  setHeaderFontWeightBold: function () {
    this.headerRowRange.setFontWeight('bold');
  },
  setFontStyle: function (style) {
    this.headerRowRange.setFontStyle(style);
  },
  addCommentToColumn: function (comment, headerRowColumnNumber) {
    var cellToComment = this.headerRowRange.getCell(1, headerRowColumnNumber);
    cellToComment.setNote(comment);
  },
  addColumnTitlesToHeaderRow: function () {
    var i,
      titleCell;
    this.spreadsheet.setNamedRange(this.headerRowRangeName, this.headerRowRange);
    for (i = 1; i <= this.columnTitleCount; i += 1) {
      titleCell = this.headerRowRange.getCell(1, i);
      if (titleCell.getValue() && !this.overwritePrevious) {
        throw {'name': 'Error',
               'message': '"HeaderRow.addColumnTitlesToHeaderRow()" Cannot overwrite previous values!'};
      }
      titleCell.setValue(this.columnTitles[i - 1]);
    }
  },
  setHeaderRowName: function (rngName) {
    this.spreadsheet.setNamedRange(rngName, this.headerRowRange);
  }
};

function TotalRow(listObject,columnFormuls, overwritePrevious) {
   var self = this;
  if (arguments.length !== 3) {
    throw {'name': 'Error',
           'message': '"TotalRow()" constructor function requires 2 arguments!'};
  }
  self.listObject = listObject;
  self.spreadsheet = spreadsheet;
  self.sheetName = sheetName;
  self.totalRowNumber = headerRowNumber;
  self.startColumnNumber = startColumnNumber;
  self.columnTitles = columnTitles;
  self.overwritePrevious = overwritePrevious;
  self.sheet = self.spreadsheet.getSheetByName(self.sheetName);
  self.columnFormulsCount = self.columnFormuls.length;
  self.totalRowRange = self.sheet.getRange(self.headerRowNumber,
                                            self.startColumnNumber,
                                            1,
                                            self.columnFormulsCount);
  self.headerRowRange.setFontWeight('normal');
  self.headerRowRange.setFontStyle('normal');
  self.addColumnTitlesToHeaderRow();
}
TotalRow.prototype = {
  constructor: HeaderRow,
//freezeHeaderRow: function () {
//    var sheet = this.sheet;
//    sheet.setFrozenRows(this.headerRowNumber);
//  },
  setHeaderFontWeightBold: function () {
    self.headerRowRange.setFontWeight('bold');
  },
  setFontStyle: function (style) {
    self.headerRowRange.setFontStyle(style);
  },
  addCommentToColumn: function (comment, headerRowColumnNumber) {
    var cellToComment = self.headerRowRange.getCell(1, headerRowColumnNumber);
    cellToComment.setNote(comment);
  },
  addColumnTitlesToHeaderRow: function () {
    var i,
      titleCell;
   self.spreadsheet.setNamedRange(self.headerRowRangeName, self.headerRowRange);
    for (i = 1; i <= this.columnTitleCount; i += 1) {
      titleCell = self.headerRowRange.getCell(1, i);
      if (titleCell.getValue() && !self.overwritePrevious) {
        throw {'name': 'Error',
               'message': '"HeaderRow.addColumnTitlesToHeaderRow()" Cannot overwrite previous values!'};
      }
      titleCell.setValue(self.columnTitles[i - 1]);
    }
  },
  setHeaderRowName: function (rngName) {
    self.spreadsheet.setNamedRange(rngName, self.headerRowRange);
  }
};

function test_HeaderRow() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetName = ss.getActiveSheet().getSheetName(),
    headerRowNumber = 3,
    startColumnNumber = 2,
    columnTitles = ['col1', 'col2', 'col3'],
    overwritePrevious = true,
    hr = new HeaderRow(ss, sheetName, headerRowNumber, startColumnNumber, columnTitles, overwritePrevious);
  hr.freezeHeaderRow();
  hr.setHeaderFontWeightBold();
  hr.setFontStyle('oblique');
  hr.addCommentToColumn('Comment added ' + Date(), 2);
  hr.setHeaderRowName('header');
}
function test_TotalRow() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetName = ss.getActiveSheet().getSheetName(),
    headerRowNumber = 3,
    startColumnNumber = 2,
    columnTitles = ['col1', 'col2', 'col3'],
    overwritePrevious = true,
    hr = new HeaderRow(ss, sheetName, headerRowNumber, startColumnNumber, columnTitles, overwritePrevious);
  tr.freezeHeaderRow();
  tr.setHeaderFontWeightBold();
  tr.setFontStyle('oblique');
  tr.addCommentToColumn('Comment added ' + Date(), 2);
  tr.setHeaderRowName('header');
}

/**
 * create a structured table from a range of data
 * @constructor ListObject
 * @param {Range} range the data range
 * @param {string} [name] the table name
 * @param {boolean} [hasHeaders=true] whether the table has headers
 * @return {ListObject} self
 */
function ListObject (range , name , hasHeaders, hasTotals) {
 var self = this;
 hasHeaders_ = fixOptional (hasHeaders , true);
 hasTotals_ =  fixOptional (hasTotals , true);
 // get the data from the range
 var range_ = range;
 var data_ = range_.getValues();
 // generate a unique name for the table if none given
 var name_ = name || 'table_'+new Date().getTime().toString(16);
 // get the header row
 var headers_ = hasHeaders_ ? data_.shift() : null;
 // get the header collection (using the VBA collection object)
 var numCols_ = range_.getNumColumns();
 self.ListColumns = new Collection();
 self.ListRows = new Collection();
 function reCalculate_ () {
  self.DataBodyRange =range_.offset(hasHeaders_ ?
  1 : 0, 0, data_.length, numCols_) ;
  self.HeaderRowRange = hasHeaders_ ?
  range_.offset( 0, 0, 1, numCols_) : null ;
  for (var i=0; i < numCols_ ; i++) {
  self.ListColumns.Add ( {
  Index:i+1 ,
  Name:hasHeaders_ ? headers_[i] : 'Column'+(i+1).toFixed(0),
  Range:hasHeaders_ ? self.HeaderRowRange.offset(0,i,1,1) : null,
  DataBodyRange:self.DataBodyRange.offset(0,i,data_.length,1)
  }, hasHeaders_ ? headers_[i] : 'Column'+(i+1).toFixed(0) );
  }
  data_.forEach(function (d,i) {
  self.ListRows.Add ( {
  Index:i+1,
  Range:self.DataBodyRange.offset(i,0,1,numCols_)
  });
  });
  }
  reCalculate_();
  return self;
 }

/**
 * Return the contiguous Range that contains the given cell.
 * Возвращает смежный диапазон Range, содержащий заданную ячейку
 * @param {String} cellA1 Location of a cell, in A1 notation.
 * @param {Sheet} sheet   (Optional) sheet to examine. Defaults
 *                          to "active" sheet.
 *
 * @return {Range} A Spreadsheet service Range object.
 */
function GetContiguousRange(cellA1,sheet) {
  // Check for parameters, handle defaults, throw error if required is missing
  if (arguments.length < 2) 
    sheet = SpreadsheetApp.getActiveSheet();
  if (arguments.length < 1)
    throw new Error("getContiguousRange(): missing required parameter.");
  
  // A "contiguous" range is a rectangular group of cells whose "edge" contains
  // cells with information, with all "past-edge" cells empty.
  // The range will be no larger than that given by "getDataRange()", so we can
  // use that range to limit our edge search.
  var fullRange = sheet.getDataRange();
  var data = fullRange.getValues();
  
  // The data array is 0-based, but spreadsheet rows & columns are 1-based.
  // We will make logic decisions based on rows & columns, and convert to
  // 0-based values to reference the data.
  var topLimit = fullRange.getRowIndex(); // always 1
  var leftLimit = fullRange.getColumnIndex(); // always 1
  var rightLimit = fullRange.getLastColumn();
  var bottomLimit = fullRange.getLastRow();
  
  // is there data in the target cell? If no, we're done.
  var contiguousRange = SpreadsheetApp.getActiveSheet().getRange(cellA1);
  var cellValue = contiguousRange.getValue();
  if (cellValue = "") return contiguousRange;
  
  // Define the limits of our starting dance floor
  var minRow = contiguousRange.getRow();
  var maxRow = minRow;
  var minCol = contiguousRange.getColumn();
  var maxCol = minCol;
  var chkCol, chkRow;  // For checking if the edge is clear

  // Now, expand our range in one direction at a time until we either reach
  // the Limits, or our next expansion would have no filled cells. Repeat
  // until no direction need expand.
  var expanding;
  do {
    expanding = false;
    // Move it to the left
    if (minCol > leftLimit) {
      chkCol = minCol - 1;
      for (var row = minRow; row <= maxRow; row++)  {
        if (data[row-1][chkCol-1] != "") {
          expanding = true;
          minCol = chkCol; // expand left 1 column
          break;
        }
      }
    }
    
    // Move it on up
    if (minRow > topLimit) {
      chkRow = minRow - 1;
      for (var col = minCol; col <= maxCol; col++)  {
        if (data[chkRow-1][col-1] != "") {
          expanding = true;
          minRow = chkRow; // expand up 1 row
          break;
        }
      }
    }
    
    // Move it to the right
    if (maxCol < rightLimit) {
      chkCol = maxCol + 1;
      for (var row = minRow; row <= maxRow; row++)  {
        if (data[row-1][chkCol-1] != "") {
          expanding = true;
          maxCol = chkCol; // expand right 1 column
          break;
        }
      }
    }
    
    // Then get on down
    if (maxRow < bottomLimit) {
      chkRow = maxRow + 1;
      for (var col = minCol; col <= maxCol; col++)  {
        if (data[chkRow-1][col-1] != "") {
          expanding = true;
          maxRow = chkRow; // expand down 1 row
          break;
        }
      }
    }
       
  } while (expanding);  // Lather, rinse, repeat
  
  // We've found the extent of our contiguous range - return a Range object.
  return sheet.getRange(minRow, minCol, (maxRow - minRow + 1), (maxCol - minCol + 1))
}

function displayInfo(args) { 
    var output = ""; 
    if (typeof args.name == "string"){ 
    output += "Name: " + args.name + "\n"; 
    } 
    if (typeof args.age == "number") { 
    output += "Age: " + args.age + "\n"; 
    } 
    alert(output); 
} 
displayInfo({ 
    name: "Nicholas", 
    age: 29 
}); 
displayInfo({ 
    name: "Greg" 
}); 

function normaliseDataTableName(){
var CurentRangeCell=GetContiguousRange("E3",sheet);
var name = DATATABLE + CurentRangeCell.getA1Notation().replace(/[^A-Z0-9]/g,"");
}