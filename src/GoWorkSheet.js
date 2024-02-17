
//options {cache,headerCellAddress, or dataRange}
const GoWorksheet = function (spreadsheet, options) {
  if (!options) options = {};
  options.cache = !!options.cache;
  this.spreadsheet = spreadsheet;
  this.worksheet = options.worksheet;

  //get range in A1Notation upto last empty row and emptycol from headerCellAddress in rectangular form
  if (options.headerCellAddress)
    this.rangeA1Notation = this.worksheet.getRange(options.headerCellAddress).getDataRegion().getA1Notation();
  else if (options.dataRange) this.rangeA1Notation = options.dataRange;
  else {
    const ws = this.worksheet;
    this.rangeA1Notation = ws.getRange(1, 1, ws.getLastRow(), ws.getLastColumn()).getDataRegion().getA1Notation();
  }

  this._gridRange = new legridRange(this.rangeA1Notation);

  if (options.cache) this._cache = this.worksheet.getRange(this.rangeA1Notation).getValues();
}

//getsheetId gid=1216109954
GoWorksheet.prototype.getId = function () {
  return this.worksheet.getSheetId();
}

//get row by rowno retuns array row[1]= header
GoWorksheet.prototype.getRow = function (rowNumber) {
  if (this._cache) return this._cache[rowNumber - 1]; //-1 becuase, 1st row header
  //return first row of the 2dArray
  const { startRowIndex, startColIndex, nosCols } = this._gridRange;
  const range = this.worksheet.getRange(startRowIndex + rowNumber - 1, startColIndex, 1, nosCols);
  return range.getValues()[0];
}
GoWorksheet.prototype.getRowAsObject = function (rowNumber) {
  var row = this.getRow(rowNumber);
  return this.rowToObject(row);
}

//getheaders of table
GoWorksheet.prototype.columnNames = function () {
  if (!this._columnNames)
    this._columnNames = this.getRow(1);
  return this._columnNames;
}

// [ 30, 'Mary', 'cat'] into { id: 30, name: 'Mary', pet: 'cat' }
GoWorksheet.prototype.rowToObject = function (row) {
  if (!row)
    return;
  var obj = {};
  for (var i = 0; i < this.columnNames().length; i++) {
    var value = row[i];
    if (typeof (value) === 'undefined') {
      obj[this.columnNames()[i]] = "";
    } else {
      obj[this.columnNames()[i]] = value;
    }
  }
  return obj;
}
// { id: 30, name: 'Mary', pet: 'cat' } into [ 30, 'Mary', 'cat']
GoWorksheet.prototype.objectToRow = function (data) {
  var row = [];
  for (var i = 0; i < this.columnNames().length; i++) {
    var value = data[this.columnNames()[i]];
    if (typeof (value) === 'undefined') {
      row.push("");
    } else {
      row.push(value);
    }
  }
  return row;
}

//all data except header
GoWorksheet.prototype.allRows = function () {

  if (this._cache) {
    if (this._cache.length === 1) return [];
    return this._cache.slice(1);
  }

  const rows = this.worksheet.getRange(this.rangeA1Notation).getValues();
  if (rows.length === 1) return [];
  return rows.slice(1);;
}

GoWorksheet.prototype.all = function () {
  const rows = this.allRows();

  const that = this;
  return rows.map(function (row) { return that.rowToObject(row) });
}

GoWorksheet.prototype.last = function () {
  var row = this._gridRange.nosRows;
  if (row === 1)
    return;
  return this.getRowAsObject(row);
}


//find region
GoWorksheet.prototype.getCol = function (colIndex) {
  let values;
  if (this._cache) {
    values = this._cache.map(function (row) { return row[colIndex]; })
    values.shift(); //remove header in cache
  } else {
    values = this.allRows().map(function (row) { return row[colIndex] });//doesnot contain header
  }
  return values;
}
GoWorksheet.prototype.indices = function (key) {
  if (!this._indices)
    this._indices = {}

  if (!this._indices[key]) {
    const colIndex = this.columnNames().indexOf(key);
    this._indices[key] = this.getCol(colIndex);
  }

  return this._indices[key];
}

//return first matching value index  from grid
GoWorksheet.prototype.rowIndex = function (key, value) {
  if (this._gridRange.nosRows == 1) return -1;

  var index = this.indices(key).indexOf(value);

  if (index === -1)
    return -1;

  return index + 2; // need to add 1 because rows start at 1 from grid, then another 1 because of headers
}

GoWorksheet.prototype.first = function (key, value) {
  var index = this.rowIndex(key, value);
  if (index == -1)
    return;
  var row = this.getRow(index);
  return this.rowToObject(row);
}

/** * Finds items in the table that match the specified value in the given key. */
GoWorksheet.prototype.find = function (key, value) {
  // Create a TextFinder to search for the value in the specified key
  const { startRowIndex, startColIndex, nosRows, nosCols } = this._gridRange;
  const finder = this.worksheet.getRange(startRowIndex, this.columnNames().indexOf(key), nosRows, nosCols).createTextFinder(value);
  const cells = finder.findAll().map(cell => cell.getRow() - 1);//.getValues();
  // Find all occurrences of the value and map their row indices
  return cells.map(index => {
    // Get the row data for the found item and convert it to an object
    const row = this.worksheet.getRange(index + 1, startColIndex, 1, nosCols).getValues()[0];
    const res = this.rowToObject(row);
    // Return the found item object with its index
    return { ...res, index: index };
  });
}
/**
 * Selects items from the table based on the provided filter object.
 */
GoWorksheet.prototype.select = function (filterObject) {
  // Initialize an array to store the selected items
  const queryItems = [];

  // Convert row data to objects for easier comparison
  const srcItems = this.all();

  // Iterate through each item in the table
  for (let i = 0; i < srcItems.length; i++) {
    let currentRow = srcItems[i];
    let matching = true;

    // Check if the current item matches the filter criteria
    for (let label in filterObject) {
      if (currentRow[label] instanceof Date) {
        // If the value is a Date object, compare timestamps
        if (currentRow[label].getTime() !== filterObject[label].getTime()) {
          matching = false;
          break;
        }
      } else {
        // Otherwise, compare values directly
        if (currentRow[label] !== filterObject[label]) {
          matching = false;
          break;
        }
      }
    }
    // If all criteria match, add the item to the selected items array
    if (matching === true) {
      queryItems.push(currentRow);
    }
  }
  return queryItems;
};

//create update
// data should be in the form {key: value, key: value} or array
GoWorksheet.prototype.append = function (data, batch) {
  var toAppend;
  if (data instanceof Array) {
    toAppend = data;
  } else {
    toAppend = this.objectToRow(data);
  }
  if (batch) {
    toAppend = data.map(row => this.objectToRow(row));//
    if (!this.appendQueue)
      this.appendQueue = []
    this.appendQueue = toAppend;
  } else {
    if (this._cache) {
      this._cache.push(toAppend);
    } else {
      var newRow = this._gridRange.startRowIndex + this._gridRange.nosRows;
      var range = this.worksheet.getRange(newRow, this._gridRange.startColIndex, 1, toAppend.length);
      range.setValues([toAppend]);
    }
    this._indices = null;
    this._lastRow = null;
  }
}

GoWorksheet.prototype.processAppends = function () {

  //this.appendQueue = this.appendQueue[0];
  if (!this.appendQueue || this.appendQueue.length == 0)
    return;

  if (this._cache) {
    for (const i = 0; i < this.appendQueue.length; i++) {
      this._cache.push(this.appendQueue[i]);
    }
    return;
  }

  //for (const i = 0; i < this.appendQueue.length; i++) {
  const newRow = this._gridRange.startRowIndex + this._gridRange.nosRows;
  const range = this.worksheet.getRange(newRow, this._gridRange.startColIndex, this.appendQueue.length, this.appendQueue[0].length);

  range.setValues(this.appendQueue);
  //}


  this._indices = null;
  this._lastRow = null;
}

// overwrites existing data, be sure to pass everything in
GoWorksheet.prototype.update = function (rowIndex, data) {
  if (rowIndex == -1) {
    Logger.log("Asking to update a row with a negative index")
    return;
  }
  let toAppend;
  if (data instanceof Array) {
    toAppend = data;
  } else {
    toAppend = this.objectToRow(data);
  }
  if (this._cache) {
    this._cache[rowIndex - 1] = toAppend;
  } else {
    const { startRowIndex, startColIndex } = this._gridRange;
    const range = this.worksheet.getRange(startRowIndex + rowIndex - 1, startColIndex, 1, toAppend.length);
    range.setValues([toAppend]);
  }
  this._indices = null;
}

//delete first row of the grid
GoWorksheet.prototype.shift = function () {
  var row = this.getRow(2);//remove 2nd row, first datarow as 1st is header
  this.removeRow(2);
  return this.rowToObject(row);
}

GoWorksheet.prototype.deleteWhere = function (callback) {

  const { nosRows } = this._gridRange;
  let rowLength = nosRows;
  // start counting at 2 because index of first row is 1, and that first row is the header
  for (var rowIndex = 2; rowIndex <= rowLength; rowIndex++) {
    var obj = this.getRowAsObject(rowIndex);
    if (callback.call(this, obj, rowIndex)) {
      this.removeRow(rowIndex);
      rowIndex--;
      rowLength--;
    }
  }
  this._lastRow = null;
  this._indices = null;
}


GoWorksheet.prototype.removeRow = function (rowNumber) {
  if (this._cache) {
    this._cache.splice(rowNumber - 1, 1);
  } else {
    const { startRowIndex } = this._gridRange;
    this.worksheet.deleteRow(startRowIndex + rowNumber - 1); //eg sri 3 +2 =5 ,but it should be 4 
  }
  this._lastRow = null;
  this._indices = null;
}


//utils

//usecase do all operations while cache is true then persist it in the worksheet 
GoWorksheet.prototype.flush = function () {
  if (!this._cache) return;
  const { startRowIndex, startColIndex } = this._gridRange;
  var range = this.worksheet.getRange(startRowIndex, startColIndex, this._cache.length, this._cache[0].length)
  range.setValues(this._cache);
}

//activate work sheet
GoWorksheet.prototype.activate = function (data) {
  this.worksheet.activate();
}

//duplicate worksheet with new name
GoWorksheet.prototype.duplicateAs = function (newName) {
  this.worksheet.activate();
  var newSheet = this.spreadsheet.duplicateActiveSheet();
  newSheet.setName(newName);
  return new GoWorksheet(this.spreadsheet, { worksheet: newSheet });
}

//copy worksheet to different spreadsheet
GoWorksheet.prototype.copyTo = function (managed_spreadsheet) {
  var copied = this.worksheet.copyTo(managed_spreadsheet.spreadsheet);
  var name = this.worksheet.getName();
  copied.setName(name);
  return new GoWorksheet(managed_spreadsheet.spreadsheet, { worksheet: copied });
}


