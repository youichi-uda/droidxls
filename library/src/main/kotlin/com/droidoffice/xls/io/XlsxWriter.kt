package com.droidoffice.xls.io

import com.droidoffice.core.ooxml.OoxmlPackage
import com.droidoffice.xls.core.CellReference
import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.DateUtil
import com.droidoffice.xls.core.Workbook
import com.droidoffice.xls.core.Worksheet
import java.io.OutputStream

/**
 * Writes a Workbook to .xlsx format (OOXML SpreadsheetML).
 */
object XlsxWriter {

    fun write(workbook: Workbook, output: OutputStream) {
        val pkg = OoxmlPackage.create()

        // Build styles and register all cell styles
        val stylesBuilder = StylesBuilder()
        for (sheet in workbook.sheets) {
            for (cell in sheet.cells()) {
                if (cell.cellStyle != null) {
                    cell.styleIndex = stylesBuilder.registerStyle(cell.cellStyle)
                }
            }
        }

        // Build shared string table
        val sharedStrings = buildSharedStrings(workbook)
        val stringIndex = sharedStrings.withIndex().associate { (i, s) -> s to i }

        // Write parts
        pkg.setPart("[Content_Types].xml", buildContentTypes(workbook).toByteArray())
        pkg.setPart("_rels/.rels", buildTopRels().toByteArray())
        pkg.setPart("xl/workbook.xml", buildWorkbook(workbook).toByteArray())
        pkg.setPart("xl/_rels/workbook.xml.rels", buildWorkbookRels(workbook).toByteArray())
        pkg.setPart("xl/styles.xml", stylesBuilder.buildStylesXml().toByteArray())

        if (sharedStrings.isNotEmpty()) {
            pkg.setPart("xl/sharedStrings.xml", buildSharedStringsXml(sharedStrings).toByteArray())
        }

        for ((index, sheet) in workbook.sheets.withIndex()) {
            val sheetXml = buildSheet(sheet, stringIndex)
            pkg.setPart("xl/worksheets/sheet${index + 1}.xml", sheetXml.toByteArray())

            // Write drawing parts for sheets with pictures
            if (sheet.pictures.isNotEmpty()) {
                DrawingWriter.writeDrawings(pkg, index, sheet.pictures)
            }
        }

        pkg.writeTo(output)
    }

    private fun buildSharedStrings(workbook: Workbook): List<String> {
        val strings = linkedSetOf<String>()
        for (sheet in workbook.sheets) {
            for (cell in sheet.cells()) {
                val v = cell.cellValue
                if (v is CellValue.Text) strings.add(v.value)
            }
        }
        return strings.toList()
    }

    private fun buildContentTypes(workbook: Workbook): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">""")
        appendLine("""  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>""")
        appendLine("""  <Default Extension="xml" ContentType="application/xml"/>""")
        appendLine("""  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>""")
        appendLine("""  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>""")
        for (i in workbook.sheets.indices) {
            appendLine("""  <Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>""")
        }
        if (workbook.sheets.any { it.cells().any { c -> c.cellValue is CellValue.Text } }) {
            appendLine("""  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>""")
        }
        // Image content types
        val imageFormats = workbook.sheets.flatMap { it.pictures }.map { it.format }.toSet()
        for (fmt in imageFormats) {
            appendLine("""  <Default Extension="${fmt.extension}" ContentType="${fmt.contentType}"/>""")
        }
        // Drawing content types
        for ((i, sheet) in workbook.sheets.withIndex()) {
            if (sheet.pictures.isNotEmpty()) {
                appendLine("""  <Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>""")
            }
        }
        appendLine("</Types>")
    }

    private fun buildTopRels(): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""")
        appendLine("""  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>""")
        appendLine("</Relationships>")
    }

    private fun buildWorkbook(workbook: Workbook): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""")
        appendLine("  <sheets>")
        for ((i, sheet) in workbook.sheets.withIndex()) {
            append("""    <sheet name="${escapeXml(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"""")
            if (sheet.isHidden) append(""" state="hidden"""")
            appendLine("/>")
        }
        appendLine("  </sheets>")
        appendLine("</workbook>")
    }

    private fun buildWorkbookRels(workbook: Workbook): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""")
        for (i in workbook.sheets.indices) {
            appendLine("""  <Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>""")
        }
        val nextId = workbook.sheets.size + 1
        appendLine("""  <Relationship Id="rId$nextId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>""")
        if (workbook.sheets.any { it.cells().any { c -> c.cellValue is CellValue.Text } }) {
            appendLine("""  <Relationship Id="rId${nextId + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>""")
        }
        appendLine("</Relationships>")
    }

    private fun buildSharedStringsXml(strings: List<String>): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.size}" uniqueCount="${strings.size}">""")
        for (s in strings) {
            appendLine("  <si><t>${escapeXml(s)}</t></si>")
        }
        appendLine("</sst>")
    }

    private fun buildSheet(sheet: Worksheet, stringIndex: Map<String, Int>): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""")

        // Freeze pane
        if (sheet.freezePane != null) {
            val fp = sheet.freezePane!!
            val topLeftCell = CellReference(fp.row, fp.col).toA1()
            appendLine("  <sheetViews><sheetView tabSelected=\"${if (sheet.isActive) 1 else 0}\" workbookViewId=\"0\">")
            appendLine("    <pane xSplit=\"${fp.col}\" ySplit=\"${fp.row}\" topLeftCell=\"$topLeftCell\" state=\"frozen\"/>")
            appendLine("  </sheetView></sheetViews>")
        }

        // Column widths and hidden columns
        val allCols = (sheet.columnWidths.keys + sheet.hiddenColumns).sorted()
        if (allCols.isNotEmpty()) {
            appendLine("  <cols>")
            for (col in allCols) {
                val colNum = col + 1
                val width = sheet.columnWidths[col] ?: 8.43
                val hidden = col in sheet.hiddenColumns
                append("    <col min=\"$colNum\" max=\"$colNum\" width=\"$width\"")
                if (sheet.columnWidths.containsKey(col)) append(" customWidth=\"1\"")
                if (hidden) append(" hidden=\"1\"")
                appendLine("/>")
            }
            appendLine("  </cols>")
        }

        // Sheet data
        appendLine("  <sheetData>")

        val cells = sheet.cells().sortedWith(compareBy({ it.reference.row }, { it.reference.col }))
        var currentRow = -1

        for (cell in cells) {
            if (cell.isEmpty) continue

            if (cell.reference.row != currentRow) {
                if (currentRow >= 0) appendLine("    </row>")
                currentRow = cell.reference.row
                append("    <row r=\"${currentRow + 1}\"")
                sheet.rowHeights[currentRow]?.let { append(" ht=\"$it\" customHeight=\"1\"") }
                if (currentRow in sheet.hiddenRows) append(" hidden=\"1\"")
                appendLine(">")
            }

            val ref = cell.reference.toA1()
            val s = if (cell.styleIndex > 0) " s=\"${cell.styleIndex}\"" else ""
            when (val v = cell.cellValue) {
                is CellValue.Text -> {
                    val idx = stringIndex[v.value]
                    if (idx != null) {
                        appendLine("      <c r=\"$ref\"$s t=\"s\"><v>$idx</v></c>")
                    } else {
                        appendLine("      <c r=\"$ref\"$s t=\"inlineStr\"><is><t>${escapeXml(v.value)}</t></is></c>")
                    }
                }
                is CellValue.Number -> appendLine("      <c r=\"$ref\"$s><v>${v.value}</v></c>")
                is CellValue.Bool -> appendLine("      <c r=\"$ref\"$s t=\"b\"><v>${if (v.value) 1 else 0}</v></c>")
                is CellValue.Formula -> {
                    append("      <c r=\"$ref\"$s")
                    if (v.cachedValue is CellValue.Text) append(" t=\"str\"")
                    appendLine(">")
                    appendLine("        <f>${escapeXml(v.expression)}</f>")
                    val cached = v.cachedValue
                    if (cached is CellValue.Number) appendLine("        <v>${cached.value}</v>")
                    else if (cached is CellValue.Text) appendLine("        <v>${escapeXml(cached.value)}</v>")
                    appendLine("      </c>")
                }
                is CellValue.Error -> appendLine("      <c r=\"$ref\"$s t=\"e\"><v>${v.code.symbol}</v></c>")
                is CellValue.DateValue -> {
                    val serial = DateUtil.dateTimeToSerial(v.value)
                    appendLine("      <c r=\"$ref\"$s><v>$serial</v></c>")
                }
                is CellValue.Empty -> { /* skip */ }
            }
        }

        if (currentRow >= 0) appendLine("    </row>")
        appendLine("  </sheetData>")

        // Sheet protection
        sheet.protection?.let { prot ->
            append("  <sheetProtection sheet=\"${if (prot.sheet) 1 else 0}\"")
            append(" objects=\"${if (prot.objects) 1 else 0}\"")
            append(" scenarios=\"${if (prot.scenarios) 1 else 0}\"")
            prot.passwordHash?.let { append(" password=\"$it\"") }
            if (prot.formatCells) append(" formatCells=\"0\"")
            if (prot.insertRows) append(" insertRows=\"0\"")
            if (prot.deleteRows) append(" deleteRows=\"0\"")
            if (prot.sort) append(" sort=\"0\"")
            if (prot.autoFilter) append(" autoFilter=\"0\"")
            appendLine("/>")
        }

        // Auto filter
        sheet.autoFilterRange?.let { range ->
            appendLine("  <autoFilter ref=\"${range.toA1()}\"/>")
        }

        // Merged cells
        if (sheet.mergedRegions.isNotEmpty()) {
            appendLine("  <mergeCells count=\"${sheet.mergedRegions.size}\">")
            for (region in sheet.mergedRegions) {
                appendLine("    <mergeCell ref=\"${region.toA1()}\"/>")
            }
            appendLine("  </mergeCells>")
        }

        // Conditional formatting
        for (cf in sheet.conditionalFormattings) {
            val sqref = cf.ranges.joinToString(" ") { it.toA1() }
            appendLine("  <conditionalFormatting sqref=\"$sqref\">")
            for (rule in cf.rules) {
                append("    <cfRule type=\"${cfTypeString(rule.type)}\" priority=\"${rule.priority}\"")
                if (rule.operator != com.droidoffice.xls.core.ConditionalOperator.NONE) {
                    append(" operator=\"${cfOperatorString(rule.operator)}\"")
                }
                appendLine(">")
                rule.formula?.let { appendLine("      <formula>${escapeXml(it)}</formula>") }
                rule.formula2?.let { appendLine("      <formula>${escapeXml(it)}</formula>") }
                appendLine("    </cfRule>")
            }
            appendLine("  </conditionalFormatting>")
        }

        // Data validations
        if (sheet.dataValidations.isNotEmpty()) {
            appendLine("  <dataValidations count=\"${sheet.dataValidations.size}\">")
            for (dv in sheet.dataValidations) {
                append("    <dataValidation type=\"${dvTypeString(dv.type)}\"")
                append(" sqref=\"${dv.range.toA1()}\"")
                if (dv.operator != com.droidoffice.xls.core.ValidationOperator.BETWEEN) {
                    append(" operator=\"${dvOperatorString(dv.operator)}\"")
                }
                if (dv.showDropDown && dv.type == com.droidoffice.xls.core.ValidationType.LIST) append(" showDropDown=\"1\"")
                if (dv.showErrorMessage) append(" showErrorMessage=\"1\"")
                if (dv.showInputMessage) append(" showInputMessage=\"1\"")
                dv.errorTitle?.let { append(" errorTitle=\"${escapeXml(it)}\"") }
                dv.errorMessage?.let { append(" error=\"${escapeXml(it)}\"") }
                dv.promptTitle?.let { append(" promptTitle=\"${escapeXml(it)}\"") }
                dv.promptMessage?.let { append(" prompt=\"${escapeXml(it)}\"") }
                appendLine(">")
                dv.formula1?.let { appendLine("      <formula1>${escapeXml(it)}</formula1>") }
                dv.formula2?.let { appendLine("      <formula2>${escapeXml(it)}</formula2>") }
                appendLine("    </dataValidation>")
            }
            appendLine("  </dataValidations>")
        }

        // Drawing reference (for pictures)
        if (sheet.pictures.isNotEmpty()) {
            appendLine("  <drawing r:id=\"rId1\"/>")
        }

        appendLine("</worksheet>")
    }

    private fun cfTypeString(type: com.droidoffice.xls.core.ConditionalType): String = when (type) {
        com.droidoffice.xls.core.ConditionalType.CELL_IS -> "cellIs"
        com.droidoffice.xls.core.ConditionalType.EXPRESSION -> "expression"
        com.droidoffice.xls.core.ConditionalType.COLOR_SCALE -> "colorScale"
        com.droidoffice.xls.core.ConditionalType.DATA_BAR -> "dataBar"
        com.droidoffice.xls.core.ConditionalType.ICON_SET -> "iconSet"
        com.droidoffice.xls.core.ConditionalType.TOP_10 -> "top10"
        com.droidoffice.xls.core.ConditionalType.ABOVE_AVERAGE -> "aboveAverage"
        com.droidoffice.xls.core.ConditionalType.DUPLICATE_VALUES -> "duplicateValues"
        com.droidoffice.xls.core.ConditionalType.UNIQUE_VALUES -> "uniqueValues"
        com.droidoffice.xls.core.ConditionalType.CONTAINS_TEXT -> "containsText"
        com.droidoffice.xls.core.ConditionalType.NOT_CONTAINS_TEXT -> "notContainsText"
        com.droidoffice.xls.core.ConditionalType.BEGINS_WITH -> "beginsWith"
        com.droidoffice.xls.core.ConditionalType.ENDS_WITH -> "endsWith"
    }

    private fun cfOperatorString(op: com.droidoffice.xls.core.ConditionalOperator): String = when (op) {
        com.droidoffice.xls.core.ConditionalOperator.BETWEEN -> "between"
        com.droidoffice.xls.core.ConditionalOperator.NOT_BETWEEN -> "notBetween"
        com.droidoffice.xls.core.ConditionalOperator.EQUAL -> "equal"
        com.droidoffice.xls.core.ConditionalOperator.NOT_EQUAL -> "notEqual"
        com.droidoffice.xls.core.ConditionalOperator.GREATER_THAN -> "greaterThan"
        com.droidoffice.xls.core.ConditionalOperator.LESS_THAN -> "lessThan"
        com.droidoffice.xls.core.ConditionalOperator.GREATER_THAN_OR_EQUAL -> "greaterThanOrEqual"
        com.droidoffice.xls.core.ConditionalOperator.LESS_THAN_OR_EQUAL -> "lessThanOrEqual"
        else -> "equal"
    }

    private fun dvTypeString(type: com.droidoffice.xls.core.ValidationType): String = when (type) {
        com.droidoffice.xls.core.ValidationType.NONE -> "none"
        com.droidoffice.xls.core.ValidationType.WHOLE -> "whole"
        com.droidoffice.xls.core.ValidationType.DECIMAL -> "decimal"
        com.droidoffice.xls.core.ValidationType.LIST -> "list"
        com.droidoffice.xls.core.ValidationType.DATE -> "date"
        com.droidoffice.xls.core.ValidationType.TIME -> "time"
        com.droidoffice.xls.core.ValidationType.TEXT_LENGTH -> "textLength"
        com.droidoffice.xls.core.ValidationType.CUSTOM -> "custom"
    }

    private fun dvOperatorString(op: com.droidoffice.xls.core.ValidationOperator): String = when (op) {
        com.droidoffice.xls.core.ValidationOperator.BETWEEN -> "between"
        com.droidoffice.xls.core.ValidationOperator.NOT_BETWEEN -> "notBetween"
        com.droidoffice.xls.core.ValidationOperator.EQUAL -> "equal"
        com.droidoffice.xls.core.ValidationOperator.NOT_EQUAL -> "notEqual"
        com.droidoffice.xls.core.ValidationOperator.GREATER_THAN -> "greaterThan"
        com.droidoffice.xls.core.ValidationOperator.LESS_THAN -> "lessThan"
        com.droidoffice.xls.core.ValidationOperator.GREATER_THAN_OR_EQUAL -> "greaterThanOrEqual"
        com.droidoffice.xls.core.ValidationOperator.LESS_THAN_OR_EQUAL -> "lessThanOrEqual"
    }

    private fun escapeXml(text: String): String = text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("'", "&apos;")
}
