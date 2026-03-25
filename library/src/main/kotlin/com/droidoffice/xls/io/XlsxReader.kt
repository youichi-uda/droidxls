package com.droidoffice.xls.io

import com.droidoffice.core.exception.InvalidFileException
import com.droidoffice.core.ooxml.OoxmlPackage
import com.droidoffice.core.ooxml.RelationshipTypes
import com.droidoffice.core.ooxml.SaxReader
import com.droidoffice.core.ooxml.parseRelationships
import com.droidoffice.xls.core.CellReference
import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.DateUtil
import com.droidoffice.xls.core.Workbook
import com.droidoffice.xls.core.Worksheet
import com.droidoffice.xls.format.CellXf
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler
import java.io.InputStream

/**
 * Reads .xlsx files (OOXML SpreadsheetML) using SAX streaming.
 */
object XlsxReader {

    fun read(input: InputStream): Workbook {
        val pkg = OoxmlPackage.open(input)
        val workbook = Workbook()

        // 1. Read styles (needed for date detection)
        pkg.getPartAsStream("xl/styles.xml")?.let { stream ->
            readStyles(stream, workbook)
        }

        // 2. Read shared strings
        pkg.getPartAsStream("xl/sharedStrings.xml")?.let { stream ->
            readSharedStrings(stream, workbook)
        }

        // 3. Read workbook.xml to get sheet names/order/state
        val sheetEntries = pkg.getPartAsStream("xl/workbook.xml")?.let { stream ->
            readWorkbookSheets(stream)
        } ?: throw InvalidFileException("Missing xl/workbook.xml")

        // 4. Read relationships to map rId → sheet file paths
        val rels = pkg.getPartAsStream("xl/_rels/workbook.xml.rels")?.let { stream ->
            parseRelationships(stream)
        } ?: emptyList()

        val rIdToTarget = rels.filter { it.type == RelationshipTypes.WORKSHEET }
            .associate { it.id to it.target }

        // 5. Read each worksheet
        for ((sheetName, rId, state) in sheetEntries) {
            val target = rIdToTarget[rId] ?: continue
            val partPath = "xl/$target"
            val sheet = workbook.addSheet(sheetName)
            sheet.isHidden = (state == "hidden" || state == "veryHidden")

            pkg.getPartAsStream(partPath)?.let { stream ->
                readWorksheet(stream, sheet, workbook)
            }
        }

        return workbook
    }

    private fun readStyles(input: InputStream, workbook: Workbook) {
        SaxReader.parse(input, object : DefaultHandler() {
            private var inNumFmts = false
            private var inCellXfs = false
            private var inCellStyleXfs = false
            private val textBuffer = StringBuilder()

            override fun startElement(uri: String, localName: String, qName: String, attributes: Attributes) {
                when (localName) {
                    "numFmts" -> inNumFmts = true
                    "cellXfs" -> inCellXfs = true
                    "cellStyleXfs" -> inCellStyleXfs = true
                    "numFmt" -> if (inNumFmts) {
                        val id = attributes.getValue("numFmtId")?.toIntOrNull() ?: return
                        val code = attributes.getValue("formatCode") ?: return
                        workbook.styleSheet.numberFormats[id] = code
                    }
                    "xf" -> if (inCellXfs && !inCellStyleXfs) {
                        val numFmtId = attributes.getValue("numFmtId")?.toIntOrNull() ?: 0
                        val fontId = attributes.getValue("fontId")?.toIntOrNull() ?: 0
                        val fillId = attributes.getValue("fillId")?.toIntOrNull() ?: 0
                        val borderId = attributes.getValue("borderId")?.toIntOrNull() ?: 0
                        workbook.styleSheet.cellXfs.add(CellXf(numFmtId, fontId, fillId, borderId))
                    }
                }
            }

            override fun endElement(uri: String, localName: String, qName: String) {
                when (localName) {
                    "numFmts" -> inNumFmts = false
                    "cellXfs" -> inCellXfs = false
                    "cellStyleXfs" -> inCellStyleXfs = false
                }
            }
        })
    }

    private fun readSharedStrings(input: InputStream, workbook: Workbook) {
        SaxReader.parse(input, object : DefaultHandler() {
            private val textBuffer = StringBuilder()
            private var inSi = false
            private var inT = false

            override fun startElement(uri: String, localName: String, qName: String, attributes: Attributes) {
                when (localName) {
                    "si" -> {
                        inSi = true
                        textBuffer.clear()
                    }
                    "t" -> inT = true
                }
            }

            override fun characters(ch: CharArray, start: Int, length: Int) {
                if (inT) textBuffer.append(ch, start, length)
            }

            override fun endElement(uri: String, localName: String, qName: String) {
                when (localName) {
                    "t" -> inT = false
                    "si" -> {
                        workbook.sharedStrings.add(textBuffer.toString())
                        inSi = false
                    }
                }
            }
        })
    }

    /**
     * Returns list of (sheetName, rId, state) from workbook.xml.
     */
    private fun readWorkbookSheets(input: InputStream): List<SheetEntry> {
        val sheets = mutableListOf<SheetEntry>()

        SaxReader.parse(input, object : DefaultHandler() {
            override fun startElement(uri: String, localName: String, qName: String, attributes: Attributes) {
                if (localName == "sheet") {
                    val name = attributes.getValue("name") ?: return
                    val rId = attributes.getValue("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
                        ?: attributes.getValue("r:id")
                        ?: return
                    val state = attributes.getValue("state") ?: "visible"
                    sheets.add(SheetEntry(name, rId, state))
                }
            }
        })

        return sheets
    }

    private data class SheetEntry(val name: String, val rId: String, val state: String)

    private fun readWorksheet(input: InputStream, sheet: Worksheet, workbook: Workbook) {
        SaxReader.parse(input, SheetHandler(sheet, workbook))
    }

    private class SheetHandler(
        private val sheet: Worksheet,
        private val workbook: Workbook,
    ) : DefaultHandler() {

        private val textBuffer = StringBuilder()
        private var currentRef: String? = null
        private var currentType: String? = null
        private var currentStyleIndex: Int = 0
        private var formulaText: String? = null
        private var valueText: String? = null

        override fun startElement(uri: String, localName: String, qName: String, attributes: Attributes) {
            textBuffer.clear()
            when (localName) {
                "c" -> {
                    currentRef = attributes.getValue("r")
                    currentType = attributes.getValue("t")
                    currentStyleIndex = attributes.getValue("s")?.toIntOrNull() ?: 0
                    formulaText = null
                    valueText = null
                }
                "row" -> {
                    val rowIndex = (attributes.getValue("r")?.toIntOrNull() ?: 1) - 1
                    val ht = attributes.getValue("ht")?.toDoubleOrNull()
                    val customHeight = attributes.getValue("customHeight") == "1"
                    if (ht != null && customHeight) {
                        sheet.rowHeights[rowIndex] = ht
                    }
                    val hidden = attributes.getValue("hidden") == "1"
                    if (hidden) sheet.hiddenRows.add(rowIndex)
                }
                "col" -> {
                    val min = (attributes.getValue("min")?.toIntOrNull() ?: 1) - 1
                    val max = (attributes.getValue("max")?.toIntOrNull() ?: 1) - 1
                    val width = attributes.getValue("width")?.toDoubleOrNull()
                    val hidden = attributes.getValue("hidden") == "1"
                    val customWidth = attributes.getValue("customWidth") == "1"
                    for (c in min..max) {
                        if (width != null && customWidth) sheet.columnWidths[c] = width
                        if (hidden) sheet.hiddenColumns.add(c)
                    }
                }
                "mergeCell" -> {
                    val ref = attributes.getValue("ref")
                    if (ref != null) {
                        try {
                            sheet.addMergedRegion(com.droidoffice.xls.core.CellRange.parse(ref))
                        } catch (_: Exception) { }
                    }
                }
                "pane" -> {
                    val state = attributes.getValue("state")
                    if (state == "frozen") {
                        val xSplit = attributes.getValue("xSplit")?.toIntOrNull() ?: 0
                        val ySplit = attributes.getValue("ySplit")?.toIntOrNull() ?: 0
                        if (xSplit > 0 || ySplit > 0) {
                            sheet.freezePane = CellReference(ySplit, xSplit)
                        }
                    }
                }
            }
        }

        override fun characters(ch: CharArray, start: Int, length: Int) {
            textBuffer.append(ch, start, length)
        }

        override fun endElement(uri: String, localName: String, qName: String) {
            when (localName) {
                "v" -> valueText = textBuffer.toString()
                "f" -> formulaText = textBuffer.toString()
                "c" -> {
                    val ref = currentRef ?: return
                    val cellRef = try {
                        CellReference.parse(ref)
                    } catch (_: Exception) {
                        return
                    }

                    val cell = sheet.cell(cellRef)
                    cell.styleIndex = currentStyleIndex
                    val rawValue = valueText ?: ""

                    if (formulaText != null) {
                        val cached = if (rawValue.isNotEmpty()) parseCellValue(rawValue, currentType, currentStyleIndex) else null
                        cell.cellValue = CellValue.Formula(formulaText!!, cached)
                    } else {
                        cell.cellValue = parseCellValue(rawValue, currentType, currentStyleIndex)
                    }

                    currentRef = null
                    currentType = null
                }
            }
        }

        private fun parseCellValue(rawValue: String, type: String?, styleIndex: Int): CellValue {
            if (rawValue.isEmpty()) return CellValue.Empty

            return when (type) {
                "s" -> {
                    val index = rawValue.toIntOrNull() ?: return CellValue.Empty
                    if (index in workbook.sharedStrings.indices) {
                        CellValue.Text(workbook.sharedStrings[index])
                    } else {
                        CellValue.Empty
                    }
                }
                "b" -> CellValue.Bool(rawValue == "1" || rawValue.equals("true", ignoreCase = true))
                "e" -> {
                    val code = com.droidoffice.xls.core.ErrorCode.entries.find { it.symbol == rawValue }
                        ?: com.droidoffice.xls.core.ErrorCode.VALUE
                    CellValue.Error(code)
                }
                "str", "inlineStr" -> CellValue.Text(rawValue)
                else -> {
                    // Number (default type) — check if it's a date
                    val num = rawValue.toDoubleOrNull()
                    if (num != null) {
                        if (workbook.styleSheet.isDateFormat(styleIndex)) {
                            CellValue.DateValue(DateUtil.serialToDateTime(num))
                        } else {
                            CellValue.Number(num)
                        }
                    } else {
                        CellValue.Text(rawValue)
                    }
                }
            }
        }
    }
}
