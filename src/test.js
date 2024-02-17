function dotest() {
  const ts = new test();
  //ts.getRowTest();
  ts.findTest();
  //ts.appendRowTest();
  //ts.appendBatchTest();
  //ts.updateRowTest();
  //ts.deleteTest();
  //ts.utilTest();
}

const test = function () {
  const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  this.ss = new GoSpreadsheet(key);
}

test.prototype.getRowTest = function () {
  const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  const ss = new GoSpreadsheet(key);
  const options = { cache: true, headerCellAddress: "B3" };
  const worksheet = ss.sheet("test", options);

  //const row = worksheet.getRow(2); //should return 1d
  //const row = worksheet.getRowAsObject(2);
  //const rowObj = worksheet.rowToObject(row);
  //const objRow = worksheet.objectToRow(rowObj);
  //const allRows2DArr= worksheet.allRows();
  //const allRowsColection= worksheet.all();
  //const colData=worksheet.getCol(3);
  //const index = worksheet.rowIndex('id', 11);
  const res = worksheet._cache;
  console.log(res);
}

test.prototype.findTest = function () {
  // const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  //const ss = new GoSpreadsheet(key);
  const options = { cache: true, headerCellAddress: "B5" }
  const worksheet = this.ss.sheet("LessorRecords", options);
  //const colData=worksheet.indices("careOf");
  //const rowindex = worksheet.rowIndex("careOf", 'SHRI BALDEV RAJ CHAWLA');//give row no in A1notation
  //const first = worksheet.first("careOf", 'SHRI BALDEV RAJ CHAWLA');
  //const find= worksheet.find("careOf", 'SHRI BALDEV RAJ CHAWLA');
  //const select = worksheet.select({ 'careOf': 'SHRI BALDEV RAJ CHAWLA', 'premiseId': '17-100' });
  //const last = worksheet.last();

  const res = worksheet._cache;
  console.log(res);
}

test.prototype.appendRowTest = function () {
  // const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  //const ss = new GoSpreadsheet(key);
  const options = { cache: false, headerCellAddress: "B3" };
  const worksheet = this.ss.sheet("test", options);

  //append test
  const dataObj = {
    'id': 10, 'Building Name': 'offic10', 'Address': 'B-1/8, Apsara Arcade, Pusa Road',
    'City': 'Karol Bagh', 'State': 'New Delhi', 'Pin Code': 110005, 'Property Type': 'Shops',
  }
  worksheet.append(dataObj);
  console.log("data appended sucessfull");
}

test.prototype.appendBatchTest = function () {
  // const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  //const ss = new GoSpreadsheet(key);
  const options = { cache: false, headerCellAddress: "B3" }
  const worksheet = this.ss.sheet("test", options);

  //append batch or multiple colomn insert
  const data = [{
    'id': 11, 'Building Name': 'KB Head Office', 'Address': 'B-1/8, Apsara Arcade, Pusa Road',
    'City': 'Karol Bagh', 'State': 'New Delhi', 'Pin Code': '110005', 'Property Type': 'Shops'
  },
  {
    'id': 12, 'Building Name': 'IAPL House', 'Address': '19, IAPL, Pusa Road',
    'City': 'Karol Bagh', 'State': 'New Delhi', 'Pin Code': '110009', 'Property Type': 'Floors'
  }];

  worksheet.append(data, true);
  worksheet.processAppends();
  console.log("Batch appended sucessfull");
}

test.prototype.updateRowTest = function () {
  // const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  //const ss = new GoSpreadsheet(key);
  const options = { cache: false, headerCellAddress: "B3" }
  const worksheet = this.ss.sheet("test", options);

  //get data and modify it
  const data = worksheet.first('id', 11);
  data['City'] = 'new delhi';

  // get rowIndex
  const rowindex = worksheet.rowIndex('id', 11);

  //update data and persists
  worksheet.update(rowindex, data);
  console.log(`Row no ${rowindex} updated`);
}

test.prototype.deleteTest = function () {
  // const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  //const ss = new GoSpreadsheet(key);
  const options = { cache: false, headerCellAddress: "B3" };
  const worksheet = this.ss.sheet("test", options);

  //delete
  //const shift = worksheet.shift();
  //worksheet.removeRow(2);
  //worksheet.deleteWhere(function (row) {return row.id <= 12;});

  const res = worksheet._cache;
  console.log(res);
}

test.prototype.utilTest = function () {
  //const key = '19j02ZEr32k5yj5dIvbM8ZcY4b7URf6v21E5F2x3e7wg';
  //const ss = new GoSpreadsheet(key);

  const options = { cache: false, headerCellAddress: "B3" };
  const worksheet = this.ss.sheet("test", options);

  //util
  //worksheet.duplicateAs('test2');
  //worksheet.flush();//persist cached value
  //const keyRentModule = "1z6pGxXdQy8-WI2PDqGhtS-mUt_9ISGzb1X5QKmTcMBw";
  //const rentSS = new GoSpreadsheet(keyRentModule);
  //worksheet.copyTo(rentSS);

  const res = worksheet._cache;
  console.log(res);
}


