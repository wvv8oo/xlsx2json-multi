// Generated by CoffeeScript 1.10.0
(function() {
  var ParserXLSX, _, _async, _events, _expat, _fs, _libxmljs, _path, _unzip, _util,
    extend = function(child, parent) { for (var key in parent) { if (hasProp.call(parent, key)) child[key] = parent[key]; } function ctor() { this.constructor = child; } ctor.prototype = parent.prototype; child.prototype = new ctor(); child.__super__ = parent.prototype; return child; },
    hasProp = {}.hasOwnProperty;

  _unzip = require('unzip');

  _fs = require('fs-extra');

  _expat = require('node-expat');

  _libxmljs = require("libxmljs");

  _ = require('underscore');

  _path = require('path');

  _async = require('async');

  _events = require('events');

  _util = require('util');

  ParserXLSX = (function(superClass) {
    extend(ParserXLSX, superClass);

    function ParserXLSX(options) {
      this.options = options;
      this.sheets = [];
    }

    ParserXLSX.prototype.extract = function(cb) {
      var stream;
      stream = _fs.createReadStream(this.options.xlsxFile).pipe(_unzip.Extract({
        path: this.options.cachePath
      }));
      stream.on('error', function(err) {
        return cb(err);
      });
      return stream.on('close', function() {
        return cb(null);
      });
    };

    ParserXLSX.prototype.parseWorkBook = function(cb) {
      var parser, self, stream, workbook;
      parser = _expat.createParser();
      self = this;
      parser.on("startElement", function(name, attrs) {
        var key;
        if (name === "sheet") {
          key = attrs["r:id"].replace(/\D+/, "sheet");
          self.sheets.push({
            name: key,
            title: attrs.name
          });
        }
      });
      parser.on("endElement", function(name) {});
      parser.on("text", function(text) {});
      parser.on("error", function(err) {
        return cb(error);
      });
      workbook = _path.join(this.options.cachePath, 'xl/workbook.xml');
      stream = _fs.createReadStream(workbook, {
        bufferSize: 64 * 1024
      });
      stream.pipe(parser);
      stream.on('error', function(err) {
        return cb(err);
      });
      return stream.on('close', function() {
        var base;
        if (typeof (base = self.options).onDidParseWorkBook === "function") {
          base.onDidParseWorkBook(self.sheets);
        }
        return cb(null);
      });
    };

    ParserXLSX.prototype.parseSharedStrings = function(cb) {
      var ssContent, ssPath;
      ssPath = _path.join(this.options.cachePath, 'xl/sharedStrings.xml');
      ssContent = _fs.readFileSync(ssPath, 'utf-8');
      this.xlsx = {
        sharedStrings: _libxmljs.parseXml(ssContent),
        namespace: {
          a: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        }
      };
      return cb(null);
    };

    ParserXLSX.prototype.columnToInt = function(col) {
      var i, letters, n;
      letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
      col = col.trim().split("");
      n = 0;
      i = 0;
      while (i < col.length) {
        n *= 26;
        n += letters.indexOf(col[i]);
        i++;
      }
      return n;
    };

    ParserXLSX.prototype.calculateDimensions = function(cells) {
      var allCols, allRows, comparator, maxCol, maxRow, minCol, minRow;
      comparator = function(a, b) {
        return a - b;
      };
      allRows = _(cells).map(function(cell) {
        return cell.row;
      }).sort(comparator);
      allCols = _(cells).map(function(cell) {
        return cell.column;
      }).sort(comparator);
      minRow = allRows[0];
      maxRow = _.last(allRows);
      minCol = allCols[0];
      maxCol = _.last(allCols);
      return [
        {
          row: minRow,
          column: minCol
        }, {
          row: maxRow,
          column: maxCol
        }
      ];
    };

    ParserXLSX.prototype.parseSheets = function(cb) {
      var index, self;
      self = this;
      index = 0;
      return _async.whilst(function() {
        return index < self.sheets.length;
      }, function(done) {
        var file, sheet, stat;
        sheet = self.sheets[index++];
        if (typeof self.options === "function" ? self.options(onShouldParseSheet(sheet)) : void 0) {
          return done(null);
        }
        file = _path.join(self.options.cachePath, 'xl', 'worksheets', sheet.name + ".xml");
        if (self.options.onWillParseSheet) {
          stat = _fs.statSync(file);
          self.options.onWillParseSheet(sheet, stat);
        }
        return self.parseSheet(file, function(err, data) {
          return self.options.onDidParseSheet(sheet, data, done);
        });
      }, cb);
    };

    ParserXLSX.prototype.sheetMergeCells = function(sheet) {
      var cell, j, len, matrix, mergeCells, result;
      result = [];
      mergeCells = sheet.find("/a:worksheet/a:mergeCells/a:mergeCell", this.xlsx.namespace);
      for (j = 0, len = mergeCells.length; j < len; j++) {
        cell = mergeCells[j];
        matrix = cell.attr('ref').value().split(':');
        result.push({
          y1: this.columnToInt(matrix[0][0]) - 1,
          x1: parseInt(matrix[0][1]) - 1,
          y2: this.columnToInt(matrix[1][0]) - 1,
          x2: parseInt(matrix[1][1]) - 1
        });
      }
      return result;
    };

    ParserXLSX.prototype.parseSheet = function(sheet, cb) {
      var Cell, CellCoords, cellNodes, cells, cols, content, d, data, error1, mergeCells, na, parseError, rows, self;
      self = this;
      data = [];
      try {
        content = _fs.readFileSync(sheet, 'utf-8');
        sheet = _libxmljs.parseXml(content);
      } catch (error1) {
        parseError = error1;
        return cb(parseError);
      }
      CellCoords = function(cell) {
        cell = cell.split(/([0-9]+)/);
        this.row = parseInt(cell[1]);
        this.column = self.columnToInt(cell[0]);
      };
      na = {
        value: function() {
          return "";
        },
        text: function() {
          return "";
        }
      };
      Cell = function(cellNode) {
        var coords, r, type, value;
        r = cellNode.attr("r").value();
        type = (cellNode.attr("t") || na).value();
        value = (cellNode.get("a:v", self.xlsx.namespace) || na).text();
        coords = new CellCoords(r);
        this.column = coords.column;
        this.row = coords.row;
        this.value = value;
        this.type = type;
      };
      mergeCells = self.sheetMergeCells(sheet);
      cellNodes = sheet.find("/a:worksheet/a:sheetData/a:row/a:c", this.xlsx.namespace);
      cells = _(cellNodes).map(function(node, index) {
        return new Cell(node);
      });
      if (cells.length === 0) {
        return cb(null, {
          merge: [],
          data: []
        });
      }
      d = sheet.get("//a:dimension/@ref", self.xlsx.namespace);
      if (d) {
        d = _.map(d.value().split(":"), function(v) {
          return new CellCoords(v);
        });
      } else {
        d = calculateDimensions(cells);
      }
      cols = d[1].column - d[0].column + 1;
      rows = d[1].row - d[0].row + 1;
      console.log(rows, 'rows');
      _(rows).times(function() {
        var _row;
        _row = [];
        _(cols).times(function() {
          _row.push("");
        });
        return data.push(_row);
      });
      _.each(cells, function(cell) {
        var i, value, values;
        value = cell.value;
        if (cell.type === "s") {
          values = self.xlsx.sharedStrings.find("//a:si[" + (parseInt(value) + 1) + "]//a:t[not(ancestor::a:rPh)]", self.xlsx.namespace);
          value = "";
          i = 0;
          while (i < values.length) {
            value += values[i].text();
            i++;
          }
        }
        if (/(\d+)\.\d{4,}/.test(value)) {
          value = String(parseFloat(value));
        }
        return data[cell.row - d[0].row][cell.column - d[0].column] = value;
      });
      console.log(data.length, 'count');
      return cb(null, {
        merge: mergeCells,
        data: data
      });
    };

    ParserXLSX.prototype.execute = function(cb) {
      var queue, self;
      self = this;
      queue = [];
      queue.push(function(done) {
        return self.extract(done);
      }, function(done) {
        return self.parseWorkBook(done);
      }, function(done) {
        return self.parseSharedStrings(done);
      }, function(done) {
        return self.parseSheets(done);
      });
      return _async.series(queue, cb);
    };

    return ParserXLSX;

  })(_events.EventEmitter);

  module.exports = ParserXLSX;

}).call(this);