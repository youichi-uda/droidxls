package com.droidoffice.xls.convert

import com.droidoffice.xls.core.CellReference
import com.droidoffice.xls.core.Worksheet

/**
 * Converts a worksheet to HTML table format.
 */
object HtmlConverter {

    fun convert(
        sheet: Worksheet,
        title: String? = null,
        includeStyle: Boolean = true,
    ): String = buildString {
        appendLine("<!DOCTYPE html>")
        appendLine("<html>")
        appendLine("<head>")
        appendLine("<meta charset=\"UTF-8\">")
        title?.let { appendLine("<title>${escapeHtml(it)}</title>") }
        if (includeStyle) {
            appendLine("""<style>
table { border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt; }
th, td { border: 1px solid #ccc; padding: 4px 8px; text-align: left; }
th { background-color: #f0f0f0; font-weight: bold; }
tr:nth-child(even) { background-color: #fafafa; }
</style>""")
        }
        appendLine("</head>")
        appendLine("<body>")

        val lastRow = sheet.lastRowIndex()
        val lastCol = sheet.lastColumnIndex()

        if (lastRow >= 0 && lastCol >= 0) {
            appendLine("<table>")

            for (row in 0..lastRow) {
                if (sheet.isRowHidden(row)) continue
                appendLine("  <tr>")
                for (col in 0..lastCol) {
                    if (sheet.isColumnHidden(col)) continue
                    val cell = sheet.getCellOrNull(CellReference(row, col))
                    val value = cell?.stringValue ?: ""
                    val tag = if (row == 0) "th" else "td"
                    appendLine("    <$tag>${escapeHtml(value)}</$tag>")
                }
                appendLine("  </tr>")
            }

            appendLine("</table>")
        }

        appendLine("</body>")
        appendLine("</html>")
    }

    private fun escapeHtml(text: String): String = text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
}
