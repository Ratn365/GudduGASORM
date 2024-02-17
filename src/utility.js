
const legridRange = function (a1notation) {
  this.init(a1notation);
}

legridRange.prototype.init = function (a1notation) {
  const rgArr = a1notation.match(/(.+):(.+$)/);
  const obj = {
    _startRowIndex: Number(rgArr[1].match(/\d+/g)[0]),
    _startColIndex: this.colNum(rgArr[1].match(/[a-z]+/gi)[0]),

    _endRowIndex: Number(rgArr[2].match(/\d+/g)[0]),
    _endColIndex: this.colNum(rgArr[2].match(/[a-z]+/gi)[0]),
  }
  this.startRowIndex = obj._startRowIndex;
  this.startColIndex = obj._startColIndex;
  this.endRowIndex = obj._endRowIndex;
  this.endColIndex = obj._endColIndex;
  this.nosRows = obj._endRowIndex - obj._startRowIndex + 1;
  this.nosCols = obj._endColIndex - obj._startColIndex + 1;

  // return obj;
}

legridRange.prototype.colNum = function (column) {
  let col = column.toUpperCase(), chr0, chr1;

  if (col.length === 1) {
    chr0 = col.charCodeAt(0) - 64;
    return chr0;
  } else if (col.length === 2) {
    chr0 = (col.charCodeAt(0) - 64) * 26;
    chr1 = col.charCodeAt(1) - 64;
    return chr0 + chr1;
  }
}
