fs = require('fs')
temp = require('temp')
path = require('path')
async = require('async')
archiver = require('archiver');

blobs = require('./blobs')

numberRegex = /^[1-9\.][\d\.]+$/

class XlsxWriter
    @write = (out, data, cb) ->
        rows = data.length
        columns = 0
        columns += 1 for key of data[0]

        writer = new XlsxWriter(out)
        writer.prepare rows, columns, (err) ->
            return cb(err) if err
            for row in data
                writer.addRow(row)
            writer.pack(cb)

    constructor: (@out) ->
        @strings = []
        @stringMap = {}
        @stringIndex = 0
        @currentRow = 0

        @haveHeader = false
        @prepared = false

        @tempPath = ''

        @sheetStream = null

        @cellMap = []
        @cellLabelMap = {}

    addRow: (obj) ->
        throw Error('Should call prepare() first!') if !@prepared

        if !@haveHeader
            @_startRow()
            col = 1
            for key of obj
                @_addCell(key, col)
                @cellMap.push(key)
                col += 1
            @_endRow()

            @haveHeader = true

        @_startRow()
        for key, col in @cellMap
            @_addCell(obj[key] || "", col + 1)
        @_endRow()

    prepare: (rows, columns, cb) ->
        # Add one extra row for the header
        dimensions = @dimensions(rows + 1, columns)

        async.series [
            (cb) => temp.mkdir 'xlsx', (err, p) =>
                @tempPath = p
                cb(err)
            (cb) => fs.mkdir(@_filename('_rels'), cb)
            (cb) => fs.mkdir(@_filename('xl'), cb)
            (cb) => fs.mkdir(@_filename('xl', '_rels'), cb)
            (cb) => fs.mkdir(@_filename('xl', 'worksheets'), cb)
            (cb) => fs.writeFile(@_filename('[Content_Types].xml'), blobs.contentTypes, cb)
            (cb) => fs.writeFile(@_filename('_rels', '.rels'), blobs.rels, cb)
            (cb) => fs.writeFile(@_filename('xl', 'workbook.xml'), blobs.workbook, cb)
            (cb) => fs.writeFile(@_filename('xl', 'styles.xml'), blobs.styles, cb)
            (cb) => fs.writeFile(@_filename('xl', '_rels', 'workbook.xml.rels'), blobs.workbookRels, cb)
            (cb) =>
                @sheetStream = fs.createWriteStream(@_filename('xl', 'worksheets', 'sheet1.xml'))
                @sheetStream.write(blobs.sheetHeader(dimensions))
                cb()
        ], (err) =>
            @prepared = true
            cb(err)

    pack: (cb) ->
        throw Error('Should call prepare() first!') if !@prepared

        zipfile = archiver('zip')
        
        output = fs.createWriteStream(@out)
        
        output.on('close', ->
          console.log('archiver has been finalized and the output file descriptor has closed.')
          cb()
        )

        zipfile.on('error', ->
          throw err;
        )

        zipfile.pipe(output);

        async.series [
            (cb) =>
                @sheetStream.write(blobs.sheetFooter)
                @sheetStream.end(cb)
            (cb) =>
                stringTable = ''
                for string in @strings
                    stringTable += blobs.string(@escapeXml(string))
                fs.writeFile(@_filename('xl', 'sharedStrings.xml'), blobs.stringsHeader(@strings.length) + stringTable + blobs.stringsFooter, cb)
            (cb) => 
                zipfile.file(@_filename('[Content_Types].xml'), {name : '[Content_Types].xml'})
                zipfile.file(@_filename('_rels', '.rels'), {name : '_rels/.rels' })
                zipfile.file(@_filename('xl', 'workbook.xml'), {name : 'xl/workbook.xml' })
                zipfile.file(@_filename('xl', 'styles.xml'), {name : 'xl/styles.xml' })
                zipfile.file(@_filename('xl', 'sharedStrings.xml'), {name : 'xl/sharedStrings.xml' })
                zipfile.file(@_filename('xl', '_rels', 'workbook.xml.rels'), {name : 'xl/_rels/workbook.xml.rels' })
                zipfile.file(@_filename('xl', 'worksheets', 'sheet1.xml'), {name : 'xl/worksheets/sheet1.xml' })
                zipfile.finalize()
        ], cb

    dimensions: (rows, columns) ->
        return "A1:" + @cell(rows, columns)

    cell: (row, col) ->
        colIndex = ''
        if @cellLabelMap[col]
            colIndex = @cellLabelMap[col]
        else
            if col == 0
                # Provide a fallback for empty spreadsheets
                row = 1
                col = 1

            input = (+col - 1).toString(26)
            while input.length
                a = input.charCodeAt(input.length - 1)
                colIndex = String.fromCharCode(a + if a >= 48 and a <= 57 then 17 else -22) + colIndex
                input = if input.length > 1 then (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) else ""
            @cellLabelMap[col] = colIndex

        return colIndex + row

    _filename: (folder, name) ->
        parts = Array::slice.call(arguments)
        parts.unshift(@tempPath)
        return path.join.apply(@, parts)

    _startRow: () ->
        @rowBuffer = blobs.startRow(@currentRow)
        @currentRow += 1

    _lookupString: (value) ->
        if !@stringMap[value]
            @stringMap[value] = @stringIndex
            @strings.push(value)
            @stringIndex += 1
        return @stringMap[value]

    _addCell: (value = '', col) ->
        row = @currentRow
        cell = @cell(row, col)

        if numberRegex.test(value)
            @rowBuffer += blobs.numberCell(value, cell)
        else
            index = @_lookupString(value)
            @rowBuffer += blobs.cell(index, cell)

    _endRow: () ->
        @sheetStream.write(@rowBuffer + blobs.endRow)

    escapeXml: (str = '') ->
        return str.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')

module.exports = XlsxWriter
