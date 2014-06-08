_XLSXParser = require('../lib').XLSXParser
_path = require 'path'

options =
  xlsxFile: _path.join __dirname, 'test.xlsx'
  cachePath: _path.join __dirname, 'cache'
  onShouldParseSheet: (sheet)->
    return sheet.name isnt 'sheet1'
  onDidParseSheet: (sheet, data, cb)->
    console.log data[0].merge
    cb null

parser = new _XLSXParser options

#发生错误
#parser.on 'error', (err)-> console.log err
#parser.on 'end', ()-> console.log 'done'
parser.execute (err)->
  return console.log err if err
  console.log '全部搞完啦'
