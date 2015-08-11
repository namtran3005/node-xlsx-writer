var XlsxWriter, archiver, async, blobs, fs, numberRegex, path, temp;

fs = require('fs');

temp = require('temp');

path = require('path');

async = require('async');

archiver = require('archiver');

blobs = require('./blobs');

numberRegex = /^[1-9\.][\d\.]+$/;

XlsxWriter = (function() {
  XlsxWriter.write = function(out, data, cb) {
    var columns, key, rows, writer;
    rows = data.length;
    columns = 0;
    for (key in data[0]) {
      columns += 1;
    }
    writer = new XlsxWriter(out);
    return writer.prepare(rows, columns, function(err) {
      var i, len, row;
      if (err) {
        return cb(err);
      }
      for (i = 0, len = data.length; i < len; i++) {
        row = data[i];
        writer.addRow(row);
      }
      return writer.pack(cb);
    });
  };

  function XlsxWriter(out1) {
    this.out = out1;
    this.strings = [];
    this.stringMap = {};
    this.stringIndex = 0;
    this.currentRow = 0;
    this.haveHeader = false;
    this.prepared = false;
    this.tempPath = '';
    this.sheetStream = null;
    this.cellMap = [];
    this.cellLabelMap = {};
  }

  XlsxWriter.prototype.addRow = function(obj) {
    var col, i, key, len, ref;
    if (!this.prepared) {
      throw Error('Should call prepare() first!');
    }
    if (!this.haveHeader) {
      this._startRow();
      col = 1;
      for (key in obj) {
        this._addCell(key, col);
        this.cellMap.push(key);
        col += 1;
      }
      this._endRow();
      this.haveHeader = true;
    }
    this._startRow();
    ref = this.cellMap;
    for (col = i = 0, len = ref.length; i < len; col = ++i) {
      key = ref[col];
      this._addCell(obj[key] || "", col + 1);
    }
    return this._endRow();
  };

  XlsxWriter.prototype.prepare = function(rows, columns, cb) {
    var dimensions;
    dimensions = this.dimensions(rows + 1, columns);
    return async.series([
      (function(_this) {
        return function(cb) {
          return temp.mkdir('xlsx', function(err, p) {
            _this.tempPath = p;
            return cb(err);
          });
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.mkdir(_this._filename('_rels'), cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.mkdir(_this._filename('xl'), cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.mkdir(_this._filename('xl', '_rels'), cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.mkdir(_this._filename('xl', 'worksheets'), cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.writeFile(_this._filename('[Content_Types].xml'), blobs.contentTypes, cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.writeFile(_this._filename('_rels', '.rels'), blobs.rels, cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.writeFile(_this._filename('xl', 'workbook.xml'), blobs.workbook, cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.writeFile(_this._filename('xl', 'styles.xml'), blobs.styles, cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          return fs.writeFile(_this._filename('xl', '_rels', 'workbook.xml.rels'), blobs.workbookRels, cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          _this.sheetStream = fs.createWriteStream(_this._filename('xl', 'worksheets', 'sheet1.xml'));
          _this.sheetStream.write(blobs.sheetHeader(dimensions));
          return cb();
        };
      })(this)
    ], (function(_this) {
      return function(err) {
        _this.prepared = true;
        return cb(err);
      };
    })(this));
  };

  XlsxWriter.prototype.pack = function(cb) {
    var output, zipfile;
    if (!this.prepared) {
      throw Error('Should call prepare() first!');
    }
    zipfile = archiver('zip');
    output = fs.createWriteStream(this.out);
    output.on('close', function() {
      console.log('archiver has been finalized and the output file descriptor has closed.');
      return cb();
    });
    zipfile.on('error', function() {
      throw err;
    });
    zipfile.pipe(output);
    return async.series([
      (function(_this) {
        return function(cb) {
          _this.sheetStream.write(blobs.sheetFooter);
          return _this.sheetStream.end(cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          var i, len, ref, string, stringTable;
          stringTable = '';
          ref = _this.strings;
          for (i = 0, len = ref.length; i < len; i++) {
            string = ref[i];
            stringTable += blobs.string(_this.escapeXml(string));
          }
          return fs.writeFile(_this._filename('xl', 'sharedStrings.xml'), blobs.stringsHeader(_this.strings.length) + stringTable + blobs.stringsFooter, cb);
        };
      })(this), (function(_this) {
        return function(cb) {
          zipfile.file(_this._filename('[Content_Types].xml'), {
            name: '[Content_Types].xml'
          });
          zipfile.file(_this._filename('_rels', '.rels'), {
            name: '_rels/.rels'
          });
          zipfile.file(_this._filename('xl', 'workbook.xml'), {
            name: 'xl/workbook.xml'
          });
          zipfile.file(_this._filename('xl', 'styles.xml'), {
            name: 'xl/styles.xml'
          });
          zipfile.file(_this._filename('xl', 'sharedStrings.xml'), {
            name: 'xl/sharedStrings.xml'
          });
          zipfile.file(_this._filename('xl', '_rels', 'workbook.xml.rels'), {
            name: 'xl/_rels/workbook.xml.rels'
          });
          zipfile.file(_this._filename('xl', 'worksheets', 'sheet1.xml'), {
            name: 'xl/worksheets/sheet1.xml'
          });
          return zipfile.finalize();
        };
      })(this)
    ], cb);
  };

  XlsxWriter.prototype.dimensions = function(rows, columns) {
    return "A1:" + this.cell(rows, columns);
  };

  XlsxWriter.prototype.cell = function(row, col) {
    var a, colIndex, input;
    colIndex = '';
    if (this.cellLabelMap[col]) {
      colIndex = this.cellLabelMap[col];
    } else {
      if (col === 0) {
        row = 1;
        col = 1;
      }
      input = (+col - 1).toString(26);
      while (input.length) {
        a = input.charCodeAt(input.length - 1);
        colIndex = String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colIndex;
        input = input.length > 1 ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) : "";
      }
      this.cellLabelMap[col] = colIndex;
    }
    return colIndex + row;
  };

  XlsxWriter.prototype._filename = function(folder, name) {
    var parts;
    parts = Array.prototype.slice.call(arguments);
    parts.unshift(this.tempPath);
    return path.join.apply(this, parts);
  };

  XlsxWriter.prototype._startRow = function() {
    this.rowBuffer = blobs.startRow(this.currentRow);
    return this.currentRow += 1;
  };

  XlsxWriter.prototype._lookupString = function(value) {
    if (!this.stringMap[value]) {
      this.stringMap[value] = this.stringIndex;
      this.strings.push(value);
      this.stringIndex += 1;
    }
    return this.stringMap[value];
  };

  XlsxWriter.prototype._addCell = function(value, col) {
    var cell, index, row;
    if (value == null) {
      value = '';
    }
    row = this.currentRow;
    cell = this.cell(row, col);
    if (numberRegex.test(value)) {
      return this.rowBuffer += blobs.numberCell(value, cell);
    } else {
      index = this._lookupString(value);
      return this.rowBuffer += blobs.cell(index, cell);
    }
  };

  XlsxWriter.prototype._endRow = function() {
    return this.sheetStream.write(this.rowBuffer + blobs.endRow);
  };

  XlsxWriter.prototype.escapeXml = function(str) {
    if (str == null) {
      str = '';
    }
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  };

  return XlsxWriter;

})();

module.exports = XlsxWriter;
