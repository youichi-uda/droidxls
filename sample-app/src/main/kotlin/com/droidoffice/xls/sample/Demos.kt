package com.droidoffice.xls.sample

import android.content.Context
import com.droidoffice.core.drawingml.BorderStyle
import com.droidoffice.core.drawingml.OfficeColor
import com.droidoffice.core.drawingml.PatternType
import com.droidoffice.xls.chart.ChartType
import com.droidoffice.xls.convert.CsvConverter
import com.droidoffice.xls.convert.HtmlConverter
import com.droidoffice.xls.core.*
import com.droidoffice.xls.drawing.ImageFormat
import com.droidoffice.xls.format.HorizontalAlignment
import com.droidoffice.xls.format.NumberFormat
import com.droidoffice.xls.format.VerticalAlignment
import com.droidoffice.xls.formula.FormulaEvaluator
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.io.File
import java.time.LocalDate

class Demos(private val context: Context) {

    private fun outputDir(): File {
        val dir = File(context.filesDir, "droidxls_samples")
        dir.mkdirs()
        return dir
    }

    // -------------------------------------------------------
    // 1. Basic Read/Write
    // -------------------------------------------------------
    suspend fun basicReadWrite(): String {
        val wb = Workbook()
        val sheet = wb.addSheet("Sheet1")

        // Various cell value types
        sheet["A1"].value = "Text"
        sheet["B1"].value = "Hello, DroidXLS!"
        sheet["A2"].value = "Number"
        sheet["B2"].value = 12345.67
        sheet["A3"].value = "Integer"
        sheet["B3"].value = 42
        sheet["A4"].value = "Boolean"
        sheet["B4"].value = true
        sheet["A5"].value = "Date"
        sheet["B5"].value = LocalDate.of(2025, 3, 25)
        sheet["A6"].value = "Japanese"
        sheet["B6"].value = "こんにちは世界"

        // Save
        val file = File(outputDir(), "01_basic.xlsx")
        file.outputStream().use { wb.save(it) }

        // Read back
        val wb2 = file.inputStream().use { Workbook.open(it) }
        val s = wb2.sheets[0]

        val sb = StringBuilder()
        sb.appendLine("Created: ${file.absolutePath} (${file.length()} bytes)")
        sb.appendLine("Sheet: ${s.name}, Cells: ${s.cells().size}")
        sb.appendLine("A1=${s["A1"].stringValue}, B1=${s["B1"].stringValue}")
        sb.appendLine("B2=${s["B2"].numericValue}, B4=${s["B4"].stringValue}")
        sb.appendLine("B6=${s["B6"].stringValue}")

        check(s["B1"].stringValue == "Hello, DroidXLS!")
        check(s["B2"].numericValue == 12345.67)
        check(s["B6"].stringValue == "こんにちは世界")

        return sb.toString()
    }

    // -------------------------------------------------------
    // 2. Styles & Formatting
    // -------------------------------------------------------
    suspend fun styles(): String {
        val wb = Workbook()
        val sheet = wb.addSheet("Styled")

        // Bold header with blue background
        sheet["A1"].value = "Product"
        sheet["B1"].value = "Price"
        sheet["C1"].value = "Stock"
        for (col in 0..2) {
            sheet.cell(0, col).style {
                font { bold = true; size = 13.0; color = OfficeColor.WHITE }
                fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.Rgb(47, 85, 151) }
                border { all = BorderStyle.THIN; allColor = OfficeColor.BLACK }
                alignment { horizontal = HorizontalAlignment.CENTER; vertical = VerticalAlignment.CENTER }
            }
        }

        // Data rows with number format
        val products = listOf("Widget" to 19.99, "Gadget" to 49.50, "Gizmo" to 99.00)
        for ((i, p) in products.withIndex()) {
            val row = i + 1
            sheet.cell(row, 0).value = p.first
            sheet.cell(row, 1).value = p.second
            sheet.cell(row, 1).style { numberFormat(NumberFormat.THOUSANDS_DECIMAL) }
            sheet.cell(row, 2).value = (i + 1) * 100.0
        }

        // Column widths
        sheet.setColumnWidth(0, 15.0)
        sheet.setColumnWidth(1, 12.0)
        sheet.setColumnWidth(2, 10.0)
        sheet.setRowHeight(0, 25.0)

        // Merged cell
        sheet["A5"].value = "Total:"
        sheet.addMergedRegion(CellRange.parse("A5:B5"))

        val file = File(outputDir(), "02_styles.xlsx")
        file.outputStream().use { wb.save(it) }

        // Read back and verify styles applied (styleIndex > 0)
        val wb2 = file.inputStream().use { Workbook.open(it) }
        val headerStyleIdx = wb2.sheets[0]["A1"].styleIndex

        return buildString {
            appendLine("Created: ${file.name} (${file.length()} bytes)")
            appendLine("Header style index: $headerStyleIdx (>0 = styled)")
            appendLine("Column A width: ${wb2.sheets[0].columnWidths[0]}")
            appendLine("Row 0 height: ${wb2.sheets[0].rowHeights[0]}")
            appendLine("Merged regions: ${wb2.sheets[0].mergedRegions}")
            check(headerStyleIdx > 0)
        }
    }

    // -------------------------------------------------------
    // 3. Formulas
    // -------------------------------------------------------
    suspend fun formulas(): String {
        val wb = Workbook()
        val sheet = wb.addSheet("Formulas")

        sheet["A1"].value = 10.0
        sheet["A2"].value = 20.0
        sheet["A3"].value = 30.0
        sheet["A4"].value = 0.0
        sheet["B1"].value = "Alice"
        sheet["B2"].value = "Bob"

        // Aggregate
        sheet["C1"].value = "SUM"
        sheet["D1"].formula = "=SUM(A1:A3)"
        sheet["C2"].value = "AVERAGE"
        sheet["D2"].formula = "=AVERAGE(A1:A3)"
        sheet["C3"].value = "MAX"
        sheet["D3"].formula = "=MAX(A1:A3)"
        sheet["C4"].value = "COUNT"
        sheet["D4"].formula = "=COUNT(A1:A4)"

        // Logic
        sheet["C5"].value = "IF"
        sheet["D5"].formula = "=IF(A1>15,\"big\",\"small\")"
        sheet["C6"].value = "IFERROR"
        sheet["D6"].formula = "=IFERROR(A1/A4,0)"

        // Math
        sheet["C7"].value = "ROUND"
        sheet["D7"].formula = "=ROUND(3.14159,2)"
        sheet["C8"].value = "ABS"
        sheet["D8"].formula = "=ABS(-42)"

        // String
        sheet["C9"].value = "CONCAT"
        sheet["D9"].formula = "=CONCATENATE(B1,\" & \",B2)"
        sheet["C10"].value = "LEN"
        sheet["D10"].formula = "=LEN(B1)"
        sheet["C11"].value = "UPPER"
        sheet["D11"].formula = "=UPPER(B1)"

        // Date
        sheet["C12"].value = "DATE"
        sheet["D12"].formula = "=DATE(2025,3,25)"

        // Recalculate
        val evaluator = FormulaEvaluator(sheet)
        evaluator.recalculateAll()

        // Verify
        val sumVal = (sheet["D1"].cellValue as CellValue.Formula).cachedValue as CellValue.Number
        val avgVal = (sheet["D2"].cellValue as CellValue.Formula).cachedValue as CellValue.Number
        val concatVal = (sheet["D9"].cellValue as CellValue.Formula).cachedValue as CellValue.Text

        val file = File(outputDir(), "03_formulas.xlsx")
        file.outputStream().use { wb.save(it) }

        return buildString {
            appendLine("Created: ${file.name}")
            appendLine("SUM(10,20,30) = ${sumVal.value}")
            appendLine("AVERAGE = ${avgVal.value}")
            appendLine("CONCATENATE = ${concatVal.value}")
            check(sumVal.value == 60.0)
            check(avgVal.value == 20.0)
            check(concatVal.value == "Alice & Bob")
        }
    }

    // -------------------------------------------------------
    // 4. Sheet Operations
    // -------------------------------------------------------
    suspend fun sheetOperations(): String {
        val wb = Workbook()

        // Add sheets
        val s1 = wb.addSheet("Data")
        s1["A1"].value = "Original"
        val s2 = wb.addSheet("Summary")
        s2["A1"].value = "Summary"

        // Copy sheet
        val s3 = wb.copySheet(0, "Data Copy")

        // Move sheet
        wb.moveSheet(2, 1)

        // Insert sheet
        wb.insertSheet(0, "Cover")
        wb.getSheet("Cover")!!["A1"].value = "Cover Page"

        // Hide a sheet
        wb.getSheet("Summary")!!.isHidden = true

        // Active sheet
        wb.activeSheetIndex = 1

        // Freeze pane
        wb.getSheet("Data")!!.freeze(1, 1)

        // Row/column operations on Data sheet
        val data = wb.getSheet("Data")!!
        data["A2"].value = "Row2"
        data["A3"].value = "Row3"
        data.insertRows(2, 1) // Insert row at index 2
        data.setColumnWidth(0, 20.0)
        data.hideColumn(1)

        val file = File(outputDir(), "04_sheet_ops.xlsx")
        file.outputStream().use { wb.save(it) }

        val wb2 = file.inputStream().use { Workbook.open(it) }

        return buildString {
            appendLine("Created: ${file.name}")
            appendLine("Sheet count: ${wb2.sheetCount}")
            appendLine("Sheets: ${wb2.sheets.map { "${it.name}${if (it.isHidden) "(hidden)" else ""}" }}")
            appendLine("Data freeze: ${wb2.getSheet("Data")?.freezePane}")
            appendLine("Data copy A1: ${wb2.getSheet("Data Copy")?.get("A1")?.stringValue}")
            check(wb2.sheetCount == 4)
            check(wb2.getSheet("Summary")?.isHidden == true)
            check(wb2.getSheet("Data Copy")?.get("A1")?.stringValue == "Original")
        }
    }

    // -------------------------------------------------------
    // 5. Data Features
    // -------------------------------------------------------
    suspend fun dataFeatures(): String {
        val wb = Workbook()
        val sheet = wb.addSheet("Data")

        sheet["A1"].value = "Name"
        sheet["B1"].value = "Score"
        sheet["C1"].value = "Grade"
        for (i in 1..5) {
            sheet.cell(i, 0).value = "Student $i"
            sheet.cell(i, 1).value = (60 + i * 8).toDouble()
        }

        // Auto filter
        sheet.autoFilterRange = CellRange.parse("A1:C6")

        // Data validation (Grade must be A-F)
        sheet.dataValidations.add(
            DataValidation(
                range = CellRange.parse("C2:C6"),
                type = ValidationType.LIST,
                formula1 = "\"A,B,C,D,F\"",
                showDropDown = true,
                showErrorMessage = true,
                errorTitle = "Invalid Grade",
                errorMessage = "Please select A, B, C, D, or F",
            )
        )

        // Conditional formatting (Score > 80 → highlight)
        sheet.conditionalFormattings.add(
            ConditionalFormatting(
                ranges = listOf(CellRange.parse("B2:B6")),
                rules = listOf(
                    ConditionalRule(
                        type = ConditionalType.CELL_IS,
                        operator = ConditionalOperator.GREATER_THAN,
                        formula = "80",
                        priority = 1,
                    )
                ),
            )
        )

        // Sheet protection
        sheet.protect("demo123")

        // Named range
        wb.namedRanges.add(NamedRange("Scores", "Data!\$B\$2:\$B\$6"))

        val file = File(outputDir(), "05_data_features.xlsx")
        file.outputStream().use { wb.save(it) }

        val wb2 = file.inputStream().use { Workbook.open(it) }

        return buildString {
            appendLine("Created: ${file.name}")
            appendLine("Auto filter: set")
            appendLine("Data validations: ${wb2.sheets[0].dataValidations.size}")
            appendLine("Conditional formatting: ${wb2.sheets[0].conditionalFormattings.size}")
            appendLine("Named ranges: ${wb.namedRanges.size}")
            appendLine("Sheet protected: ${sheet.protection != null}")
        }
    }

    // -------------------------------------------------------
    // 6. Password Protection
    // -------------------------------------------------------
    suspend fun passwordProtection(): String {
        val wb = Workbook()
        val sheet = wb.addSheet("Secret")
        sheet["A1"].value = "Confidential Data"
        sheet["A2"].value = "SSN: 123-45-6789"
        sheet["A3"].value = "Salary: $120,000"

        // Save with password
        val buffer = ByteArrayOutputStream()
        wb.save(buffer, "s3cret!")
        val encrypted = buffer.toByteArray()

        // Verify can't open without password
        var blocked = false
        try {
            Workbook.open(ByteArrayInputStream(encrypted))
        } catch (e: Exception) {
            blocked = true
        }

        // Open with correct password
        val wb2 = Workbook.open(ByteArrayInputStream(encrypted), "s3cret!")

        // Save decrypted copy
        val file = File(outputDir(), "06_password.xlsx")
        file.outputStream().use { wb2.save(it) }

        return buildString {
            appendLine("Encrypted size: ${encrypted.size} bytes")
            appendLine("Blocked without password: $blocked")
            appendLine("Decrypted A1: ${wb2.sheets[0]["A1"].stringValue}")
            appendLine("Decrypted A3: ${wb2.sheets[0]["A3"].stringValue}")
            check(blocked)
            check(wb2.sheets[0]["A1"].stringValue == "Confidential Data")
        }
    }

    // -------------------------------------------------------
    // 7. CSV / HTML Export
    // -------------------------------------------------------
    suspend fun csvHtmlExport(): String {
        val wb = Workbook()
        val sheet = wb.addSheet("Export")
        sheet["A1"].value = "Name"
        sheet["B1"].value = "City"
        sheet["C1"].value = "Score"
        sheet["A2"].value = "Alice"
        sheet["B2"].value = "Tokyo"
        sheet["C2"].value = 95.0
        sheet["A3"].value = "Bob"
        sheet["B3"].value = "Osaka"
        sheet["C3"].value = 87.0
        sheet["A4"].value = "Carol, Jr."
        sheet["B4"].value = "Kyoto"
        sheet["C4"].value = 92.0

        // CSV
        val csv = CsvConverter.convertToString(sheet)
        val csvFile = File(outputDir(), "07_export.csv")
        csvFile.writeText(csv)

        // HTML
        val html = HtmlConverter.convert(sheet, title = "Score Report")
        val htmlFile = File(outputDir(), "07_export.html")
        htmlFile.writeText(html)

        return buildString {
            appendLine("CSV (${csvFile.length()} bytes):")
            appendLine(csv.take(200))
            appendLine("HTML (${htmlFile.length()} bytes): saved")
            appendLine("HTML contains <table>: ${html.contains("<table>")}")
            appendLine("CSV escapes comma: ${csv.contains("\"Carol, Jr.\"")}")
            check(csv.contains("\"Carol, Jr.\""))
            check(html.contains("<th>Name</th>"))
        }
    }

    // -------------------------------------------------------
    // 8. Full Report (All Features Combined)
    // -------------------------------------------------------
    suspend fun fullReport(): String {
        val wb = Workbook()

        // --- Cover sheet ---
        val cover = wb.addSheet("Cover")
        cover["A1"].value = "DroidXLS Demo Report"
        cover["A1"].style {
            font { bold = true; size = 20.0; color = OfficeColor.Rgb(0, 51, 102) }
        }
        cover.addMergedRegion(CellRange.parse("A1:E1"))
        cover["A3"].value = "Generated by DroidXLS Sample App"
        cover["A4"].value = "Date: ${LocalDate.now()}"

        // --- Sales Data ---
        val sales = wb.addSheet("Sales")
        val headers = listOf("Product", "Q1", "Q2", "Q3", "Q4", "Total")
        for ((col, h) in headers.withIndex()) {
            sales.cell(0, col).value = h
            sales.cell(0, col).style {
                font { bold = true; color = OfficeColor.WHITE }
                fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.Rgb(47, 85, 151) }
                border { all = BorderStyle.THIN }
                alignment { horizontal = HorizontalAlignment.CENTER }
            }
        }

        val data = listOf(
            listOf("Widget A", 12500, 15300, 14200, 18900),
            listOf("Widget B", 8700, 9100, 11500, 12300),
            listOf("Service X", 25000, 28000, 32000, 35000),
        )
        for ((row, d) in data.withIndex()) {
            val r = row + 1
            sales.cell(r, 0).value = d[0] as String
            for (q in 1..4) {
                sales.cell(r, q).value = (d[q] as Int).toDouble()
                sales.cell(r, q).style { numberFormat(NumberFormat.THOUSANDS) }
            }
            sales.cell(r, 5).formula = "=SUM(B${r + 1}:E${r + 1})"
            sales.cell(r, 5).style {
                font { bold = true }
                numberFormat(NumberFormat.THOUSANDS)
            }
        }

        // Totals row
        val totalRow = data.size + 1
        sales.cell(totalRow, 0).value = "TOTAL"
        sales.cell(totalRow, 0).style { font { bold = true } }
        for (col in 1..5) {
            val colL = CellReference.columnToLetters(col)
            sales.cell(totalRow, col).formula = "=SUM(${colL}2:${colL}${totalRow})"
            sales.cell(totalRow, col).style {
                font { bold = true }
                numberFormat(NumberFormat.THOUSANDS)
                border {
                    top = com.droidoffice.xls.format.CellBorder(BorderStyle.DOUBLE, OfficeColor.BLACK)
                }
            }
        }

        sales.freeze(1, 0)
        sales.autoFilterRange = CellRange.parse("A1:F${totalRow + 1}")
        for (col in 0..5) sales.setColumnWidth(col, if (col == 0) 15.0 else 12.0)

        // Chart
        sales.addChart(ChartType.COLUMN) {
            title = "Quarterly Sales"
            series("Sales!\$B\$2:\$B\$4") { name = "Q1" }
            series("Sales!\$C\$2:\$C\$4") { name = "Q2" }
        }

        // Recalculate
        FormulaEvaluator(sales).recalculateAll()

        // --- Image demo (fake 1x1 PNG) ---
        val fakePng = byteArrayOf(
            0x89.toByte(), 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
        )
        cover.addPicture(fakePng, ImageFormat.PNG, 4, 0, 6, 3, 120, 60)

        // Save
        val file = File(outputDir(), "08_full_report.xlsx")
        file.outputStream().use { wb.save(it) }

        // Read back
        val wb2 = file.inputStream().use { Workbook.open(it) }

        // Verify
        val salesSheet = wb2.getSheet("Sales")!!
        val sumFormula = salesSheet["F2"].formula
        val totalQ1 = (salesSheet.cell(totalRow, 1).cellValue as? CellValue.Formula)

        // Export CSV of sales
        val csv = CsvConverter.convertToString(salesSheet)
        File(outputDir(), "08_sales.csv").writeText(csv)

        // Export HTML
        val html = HtmlConverter.convert(salesSheet, title = "Sales Report")
        File(outputDir(), "08_sales.html").writeText(html)

        return buildString {
            appendLine("Created: ${file.name} (${file.length()} bytes)")
            appendLine("Sheets: ${wb2.sheets.map { it.name }}")
            appendLine("Sales F2 formula: $sumFormula")
            appendLine("Freeze pane: ${salesSheet.freezePane}")
            appendLine("Charts: ${sales.charts.size}")
            appendLine("Pictures: ${cover.pictures.size}")
            appendLine("CSV exported, HTML exported")
            check(wb2.sheetCount == 2)
            check(sumFormula != null && sumFormula.contains("SUM"))
            check(salesSheet.freezePane == CellReference(1, 0))
        }
    }
}
