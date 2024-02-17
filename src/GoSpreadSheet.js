
//https://github.com/jsoma/gs-spreadsheet-manager

const GoSpreadsheet = function (key) {
  this.spreadsheet = key ? SpreadsheetApp.openById(key) : SpreadsheetApp.getActiveSpreadsheet();
  this.worksheets = {};
}

GoSpreadsheet.prototype.sheet = function (name, options) {
  if (!options) options = {};

  if (!this.worksheets[name]) {
    var worksheet = this.spreadsheet.getSheetByName(name);
    if (!worksheet) return;

    options.worksheet = worksheet;
    this.worksheets[name] = new GoWorksheet(this.spreadsheet, options);
  }
  return this.worksheets[name];
}

GoSpreadsheet.prototype.getId = function () {
  return this.spreadsheet.getId();
}

GoSpreadsheet.prototype.sheetCount = function () {
  return this.spreadsheet.getSheets().length;
}
//call processapend of goWorksheet for batch addition
GoSpreadsheet.prototype.processAppends = function () {
  for (var key in this.worksheets) {
    this.worksheets[key].processAppends();
  };
}

GoSpreadsheet.prototype.atomic = function (action) {
  // Copy the spreadsheet into a new spreadsheet
  this.fork();
  // Run operations
  action.call(this);
  // Commit cached values to the spreadsheet
  this.flush();
  // Move the worksheets back into the initial spreadsheet
  this.merge();
}

//creates copy spreadsheet
GoSpreadsheet.prototype.fork = function () {
  this.original_spreadsheet = this.spreadsheet;
  this.spreadsheet = this.spreadsheet.copy("Temporary " + this.spreadsheet.getName() + ", " + new Date());
  this.worksheets = {};
}

//call flush of goWorksheet for batch addition
GoSpreadsheet.prototype.flush = function () {
  for (var key in this.worksheets) {
    this.worksheets[key].flush();
  }
}

GoSpreadsheet.prototype.merge = function () {
  // Can't have a spreadsheet without sheets, so we'll add in a placeholder
  var placeholder = this.original_spreadsheet.insertSheet("Placeholder: " + new Date());

  var that = this;

  // First we'll remove every sheet except the placeholder sheet
  this.original_spreadsheet.getSheets().forEach(function (sheet) {
    if (sheet.getSheetId() === placeholder.getSheetId())
      return;

    sheet.activate();
    that.original_spreadsheet.deleteActiveSheet();
  });

  this.spreadsheet.getSheets().forEach(function (sheet) {
    var new_sheet = sheet.copyTo(that.original_spreadsheet);
    new_sheet.setName(sheet.getName());
  })

  // Remove the placeholder sheet
  placeholder.activate();
  this.original_spreadsheet.deleteActiveSheet();

  // Remove the copied document
  var file = DriveApp.getFileById(this.spreadsheet.getId());
  file.setTrashed(true);

  // Go back to the original spreadsheet
  this.spreadsheet = this.original_spreadsheet;
  this.worksheets = {};
}

GoSpreadsheet.prototype.moveFileToFolder = function (folderId) {

  // Get the file and folder by their IDs
  var file = DriveApp.getFileById(this.spreadsheet.getId());
  var folder = DriveApp.getFolderById(folderId);

  try {
    // Move the file to the folder
    file.moveTo(folder);
    // folder.createFile(file.getBlob());
    file.setTrashed(true);

    Logger.log('File moved successfully!');
  } catch (error) {
    Logger.log('Error moving file: ' + error.toString());
  }
}

