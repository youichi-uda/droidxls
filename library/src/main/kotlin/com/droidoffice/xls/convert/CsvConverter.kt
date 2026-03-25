package com.droidoffice.xls.convert

import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.Worksheet
import java.io.OutputStream
import java.io.Writer

/**
 * Converts a worksheet to CSV format.
 */
object CsvConverter {

    fun convert(
        sheet: Worksheet,
        output: OutputStream,
        delimiter: Char = ',',
        encoding: String = "UTF-8",
        includeHeaders: Boolean = true,
    ) {
        output.writer(charset(encoding)).use { writer ->
            convert(sheet, writer, delimiter)
        }
    }

    fun convert(
        sheet: Worksheet,
        writer: Writer,
        delimiter: Char = ',',
    ) {
        val lastRow = sheet.lastRowIndex()
        val lastCol = sheet.lastColumnIndex()
        if (lastRow < 0 || lastCol < 0) return

        for (row in 0..lastRow) {
            for (col in 0..lastCol) {
                if (col > 0) writer.write(delimiter.toString())
                val cell = sheet.getCellOrNull(com.droidoffice.xls.core.CellReference(row, col))
                if (cell != null) {
                    writer.write(escapeCsv(cell.stringValue, delimiter))
                }
            }
            writer.write("\r\n")
        }
        writer.flush()
    }

    fun convertToString(sheet: Worksheet, delimiter: Char = ','): String {
        val sb = StringBuilder()
        val lastRow = sheet.lastRowIndex()
        val lastCol = sheet.lastColumnIndex()
        if (lastRow < 0 || lastCol < 0) return ""

        for (row in 0..lastRow) {
            for (col in 0..lastCol) {
                if (col > 0) sb.append(delimiter)
                val cell = sheet.getCellOrNull(com.droidoffice.xls.core.CellReference(row, col))
                if (cell != null) {
                    sb.append(escapeCsv(cell.stringValue, delimiter))
                }
            }
            sb.append("\r\n")
        }
        return sb.toString()
    }

    private fun escapeCsv(value: String, delimiter: Char): String {
        if (value.isEmpty()) return ""
        val needsQuoting = value.contains(delimiter) || value.contains('"') ||
            value.contains('\n') || value.contains('\r')
        return if (needsQuoting) {
            "\"${value.replace("\"", "\"\"")}\""
        } else {
            value
        }
    }
}
