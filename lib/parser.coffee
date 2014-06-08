_unzip = require 'unzip'
_fs = require 'fs-extra'
_expat = require 'node-expat'
_libxmljs = require "libxmljs"
_ = require 'underscore'
_path = require 'path'
_async = require 'async'
_events = require 'events'
_util = require 'util'

class ParserXLSX extends _events.EventEmitter
  constructor: (@options)->
    #_events.EventEmitter.call @
    @sheets = []

  #释放文件到目标目录
  extract: (cb)->
    stream = _fs.createReadStream(@options.xlsxFile).pipe _unzip.Extract(path: @options.cachePath)
    stream.on 'error', (err)-> cb err
    stream.on 'close', ()-> cb null

  #分析总目录
  parseWorkBook: (cb)->
    parser = _expat.createParser()
    self = @

    parser.on "startElement", (name, attrs) ->
      if name is "sheet"
        key = attrs["r:id"].replace(/\D+/, "sheet")
        self.sheets.push
          name: key
          title: attrs.name
      return

    parser.on "endElement", (name) ->
    parser.on "text", (text) ->
    parser.on "error", (err) -> cb error

    workbook = _path.join @options.cachePath, 'xl/workbook.xml'
    stream = _fs.createReadStream(workbook, bufferSize: 64 * 1024)
    stream.pipe parser
    stream.on 'error', (err)-> cb err
    stream.on 'close', ->
      self.options.onDidParseWorkBook?(self.sheets)
      cb null

  #转换所有
  parseSharedStrings: (cb)->
    #读取sharedString
    ssPath = _path.join @options.cachePath, 'xl/sharedStrings.xml'
    ssContent = _fs.readFileSync ssPath, 'utf-8'
    @xlsx =
      sharedStrings: _libxmljs.parseXml ssContent
      namespace: a: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    cb null

  #转换列名到int索引
  columnToInt: (col)->
    letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    col = col.trim().split("")
    n = 0
    i = 0

    while i < col.length
      n *= 26
      n += letters.indexOf(col[i])
      i++
    n

  calculateDimensions: (cells) ->
    comparator = (a, b) ->
      a - b

    allRows = _(cells).map((cell) ->
      cell.row
    ).sort(comparator)
    allCols = _(cells).map((cell) ->
      cell.column
    ).sort(comparator)
    minRow = allRows[0]
    maxRow = _.last(allRows)
    minCol = allCols[0]
    maxCol = _.last(allCols)
    [
      {
        row: minRow
        column: minCol
      }
      {
        row: maxRow
        column: maxCol
      }
    ]

  #转换多个sheets
  parseSheets: (cb)->
    self = @
    index = 0
    _async.whilst(
      -> index < self.sheets.length
      (done)->
        sheet = self.sheets[index++]
        #检查是否需要跳过
        return done null if self.options?onShouldParseSheet sheet
        file = _path.join(self.options.cachePath, 'xl', 'worksheets', "#{sheet.name}.xml")
        self.parseSheet file, (err, data)->
          self.options.onDidParseSheet sheet, data, done
      cb
    )

  #分析具体的sheet，摘自第三方代码，暂未统一风格
  parseSheet: (sheet, cb)->
    self = @
    data = []
    try
      content = _fs.readFileSync sheet, 'utf-8'
      sheet = _libxmljs.parseXml content
    catch parseError
      return cb parseError

    CellCoords = (cell) ->
      cell = cell.split(/([0-9]+)/)
      @row = parseInt(cell[1])
      @column = self.columnToInt(cell[0])
      return

    na =
      value: -> ""
      text: -> ""

    Cell = (cellNode) ->
      r = cellNode.attr("r").value()
      type = (cellNode.attr("t") or na).value()
      value = (cellNode.get("a:v", self.xlsx.namespace) or na).text()
      coords = new CellCoords(r)
      @column = coords.column
      @row = coords.row
      @value = value
      @type = type
      return

    cellNodes = sheet.find("/a:worksheet/a:sheetData/a:row/a:c", @xlsx.namespace)
    cells = _(cellNodes).map((node) ->
      new Cell(node)
    )
    d = sheet.get("//a:dimension/@ref", self.xlsx.namespace)
    if d
      d = _.map(d.value().split(":"), (v) ->
        new CellCoords(v)
      )
    else
      d = calculateDimensions(cells)

    cols = d[1].column - d[0].column + 1
    rows = d[1].row - d[0].row + 1
    _(rows).times ->
      _row = []
      _(cols).times ->
        _row.push ""
        return

      data.push _row
      #return

    merge = []
    mergeCell = null
    lastRowIndex = 0
    _.each cells, (cell) ->
      value = cell.value
      #推断为合并单元格
      if not cell.type and not cell.value
        #已经有merge了
        if mergeCell
          mergeCell.colspan++
        else
          mergeCell = colspan: 2, start: cell.column - 2
      else
        #存在mergeCell，则将mergeCell加入到列表中，
        if mergeCell
          merge.push _.extend {}, mergeCell
          mergeCell = null

      if cell.type is "s"
        values = self.xlsx.sharedStrings.find("//a:si[" + (parseInt(value) + 1) + "]//a:t[not(ancestor::a:rPh)]", self.xlsx.namespace)
        value = ""
        i = 0

        while i < values.length
          value += values[i].text()
          i++

      value = String(parseFloat(value)) if /(\d+)\.\d{4,}/.test value
      rowIndex = cell.row - d[0].row
      row = data[rowIndex]
      row[cell.column - d[0].column] = value

      #新的一行了
      if lastRowIndex != rowIndex and merge.length > 0
        data[rowIndex - 1].merge = merge
        merge = []

      lastRowIndex = rowIndex

    #data[rowIndex].merge = merge if merge.length > 0
    cb null, data

  #执行操作
  execute: (cb)->
    self = @
    queue = []
    queue.push(
      #释放出sheets
      (done)-> self.extract done
      (done)-> self.parseWorkBook done
      (done)-> self.parseSharedStrings done
      (done)-> self.parseSheets done
    )

    _async.series queue, cb


module.exports = ParserXLSX