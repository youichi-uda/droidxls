package com.droidoffice.xls.e2e

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
import com.droidoffice.xls.formula.FormulaEvaluator
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.time.LocalDate
import kotlin.test.assertEquals
import kotlin.test.assertNotNull
import kotlin.test.assertTrue

/**
 * End-to-end tests: full realistic workflows from creation to export.
 */
class FullWorkflowTest {

    @Test
    fun `sales report workflow - create, style, formula, save, reopen, export`() {
        // === 1. Create workbook with realistic data ===
        val wb = Workbook()
        val sheet = wb.addSheet("Sales Report")

        // Headers
        val headers = listOf("Product", "Q1", "Q2", "Q3", "Q4", "Total")
        for ((col, header) in headers.withIndex()) {
            sheet.cell(0, col).value = header
            sheet.cell(0, col).style {
                font { bold = true; size = 12.0; color = OfficeColor.WHITE }
                fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.Rgb(47, 85, 151) }
                border { all = BorderStyle.THIN; allColor = OfficeColor.BLACK }
                alignment { horizontal = HorizontalAlignment.CENTER }
            }
        }

        // Data rows
        val products = listOf("Widget A", "Widget B", "Widget C", "Service X", "Service Y")
        val data = listOf(
            listOf(12500.0, 15300.0, 14200.0, 18900.0),
            listOf(8700.0, 9100.0, 11500.0, 12300.0),
            listOf(3200.0, 4500.0, 3800.0, 5100.0),
            listOf(25000.0, 28000.0, 32000.0, 35000.0),
            listOf(15000.0, 17500.0, 19000.0, 21500.0),
        )

        for ((row, product) in products.withIndex()) {
            val r = row + 1
            sheet.cell(r, 0).value = product
            for ((col, value) in data[row].withIndex()) {
                sheet.cell(r, col + 1).value = value
                sheet.cell(r, col + 1).style {
                    numberFormat(NumberFormat.THOUSANDS_DECIMAL)
                }
            }
            // Total formula
            val lastDataCol = CellReference.columnToLetters(4)
            val firstDataCol = CellReference.columnToLetters(1)
            sheet.cell(r, 5).formula = "=SUM(${firstDataCol}${r + 1}:${lastDataCol}${r + 1})"
        }

        // Grand total row
        val totalRow = products.size + 1
        sheet.cell(totalRow, 0).value = "TOTAL"
        sheet.cell(totalRow, 0).style { font { bold = true } }
        for (col in 1..5) {
            val colLetter = CellReference.columnToLetters(col)
            sheet.cell(totalRow, col).formula = "=SUM(${colLetter}2:${colLetter}${totalRow})"
            sheet.cell(totalRow, col).style {
                font { bold = true }
                border { top = com.droidoffice.xls.format.CellBorder(BorderStyle.DOUBLE, OfficeColor.BLACK) }
            }
        }

        // Column widths
        sheet.setColumnWidth(0, 15.0)
        for (col in 1..5) sheet.setColumnWidth(col, 12.0)

        // Freeze header row
        sheet.freeze(1, 0)

        // Auto filter
        sheet.autoFilterRange = CellRange.parse("A1:F${totalRow + 1}")

        // === 2. Recalculate formulas ===
        val evaluator = FormulaEvaluator(sheet)
        evaluator.recalculateAll()

        // Verify formula results
        val totalQ1 = sheet.cell(totalRow, 1).let {
            (it.cellValue as CellValue.Formula).cachedValue as CellValue.Number
        }.value
        assertEquals(64400.0, totalQ1) // 12500+8700+3200+25000+15000

        // === 3. Add a second sheet with chart data ===
        val summarySheet = wb.addSheet("Summary")
        summarySheet["A1"].value = "Total Revenue"
        summarySheet["A2"].value = "Product Count"
        summarySheet["B1"].value = 0.0  // Would be calculated
        summarySheet["B2"].value = products.size.toDouble()

        // Add chart
        summarySheet.addChart(ChartType.BAR) {
            title = "Quarterly Sales"
            fromCol = 0; fromRow = 4; toCol = 8; toRow = 18
            series("'Sales Report'!\$B\$2:\$B\$6") { name = "Q1" }
            series("'Sales Report'!\$C\$2:\$C\$6") { name = "Q2" }
        }

        // === 4. Save to xlsx ===
        val xlsxBuffer = ByteArrayOutputStream()
        wb.save(xlsxBuffer)
        val xlsxBytes = xlsxBuffer.toByteArray()
        assertTrue(xlsxBytes.size > 0, "xlsx output should not be empty")

        // === 5. Reopen and verify ===
        val wb2 = Workbook.open(ByteArrayInputStream(xlsxBytes))
        assertEquals(2, wb2.sheetCount)
        assertEquals("Sales Report", wb2.sheets[0].name)
        assertEquals("Summary", wb2.sheets[1].name)

        val reopened = wb2.sheets[0]

        // Verify headers
        assertEquals("Product", reopened["A1"].stringValue)
        assertEquals("Total", reopened["F1"].stringValue)

        // Verify data
        assertEquals("Widget A", reopened["A2"].stringValue)
        assertEquals(12500.0, reopened["B2"].numericValue)

        // Verify formulas preserved
        assertNotNull(reopened["F2"].formula, "Formula should be preserved")
        assertTrue(reopened["F2"].formula!!.contains("SUM"))

        // Verify freeze pane
        assertEquals(CellReference(1, 0), reopened.freezePane)

        // Verify column widths
        assertEquals(15.0, reopened.columnWidths[0])

        // === 6. Export to CSV ===
        val csv = CsvConverter.convertToString(reopened)
        assertTrue(csv.contains("Product"))
        assertTrue(csv.contains("Widget A"))
        assertTrue(csv.contains("12500"))

        // === 7. Export to HTML ===
        val html = HtmlConverter.convert(reopened, title = "Sales Report")
        assertTrue(html.contains("<title>Sales Report</title>"))
        assertTrue(html.contains("Widget A"))
        assertTrue(html.contains("<th>Product</th>"))
    }

    @Test
    fun `multi-feature workbook - styles, images, validation, protection, merged cells`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Invoice")

        // Merged header
        sheet["A1"].value = "INVOICE"
        sheet["A1"].style {
            font { bold = true; size = 24.0; color = OfficeColor.Rgb(0, 51, 102) }
            alignment { horizontal = HorizontalAlignment.CENTER }
        }
        sheet.addMergedRegion(CellRange.parse("A1:F1"))

        // Company info
        sheet["A3"].value = "Bill To:"
        sheet["A3"].style { font { bold = true } }
        sheet["A4"].value = "Acme Corporation"
        sheet["A5"].value = "123 Business St"
        sheet["A6"].value = "Tokyo, Japan"

        // Invoice details
        sheet["D3"].value = "Invoice #"
        sheet["E3"].value = "INV-2025-001"
        sheet["D4"].value = "Date"
        sheet["E4"].cellValue = CellValue.DateValue(LocalDate.of(2025, 3, 25).atStartOfDay())
        sheet["D5"].value = "Due Date"
        sheet["E5"].cellValue = CellValue.DateValue(LocalDate.of(2025, 4, 25).atStartOfDay())

        // Line items table
        val tableHeaders = listOf("Item", "Description", "Qty", "Unit Price", "Amount")
        for ((col, h) in tableHeaders.withIndex()) {
            sheet.cell(8, col).value = h
            sheet.cell(8, col).style {
                font { bold = true; color = OfficeColor.WHITE }
                fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.Rgb(0, 51, 102) }
            }
        }

        // Line items
        val items = listOf(
            Triple("DroidXLS License", "Annual commercial license", Pair(1, 99.0)),
            Triple("Support Package", "Priority email support", Pair(1, 199.0)),
            Triple("Custom Integration", "API integration consulting", Pair(5, 150.0)),
        )
        for ((i, item) in items.withIndex()) {
            val row = 9 + i
            sheet.cell(row, 0).value = item.first
            sheet.cell(row, 1).value = item.second
            sheet.cell(row, 2).value = item.third.first.toDouble()
            sheet.cell(row, 3).value = item.third.second
            sheet.cell(row, 3).style { numberFormat(NumberFormat.THOUSANDS_DECIMAL) }
            val colC = CellReference.columnToLetters(2)
            val colD = CellReference.columnToLetters(3)
            sheet.cell(row, 4).formula = "=${colC}${row + 1}*${colD}${row + 1}"
            sheet.cell(row, 4).style { numberFormat(NumberFormat.THOUSANDS_DECIMAL) }
        }

        // Total
        val totalRow = 9 + items.size
        sheet.cell(totalRow, 3).value = "TOTAL:"
        sheet.cell(totalRow, 3).style { font { bold = true }; alignment { horizontal = HorizontalAlignment.RIGHT } }
        val colE = CellReference.columnToLetters(4)
        sheet.cell(totalRow, 4).formula = "=SUM(${colE}10:${colE}${totalRow})"
        sheet.cell(totalRow, 4).style {
            font { bold = true }
            numberFormat(NumberFormat.THOUSANDS_DECIMAL)
            border { top = com.droidoffice.xls.format.CellBorder(BorderStyle.DOUBLE, OfficeColor.BLACK) }
        }

        // Data validation on Qty column (must be positive integer)
        sheet.dataValidations.add(
            DataValidation(
                range = CellRange.parse("C10:C${totalRow - 1}"),
                type = ValidationType.WHOLE,
                operator = ValidationOperator.GREATER_THAN,
                formula1 = "0",
                showErrorMessage = true,
                errorTitle = "Invalid Quantity",
                errorMessage = "Quantity must be a positive integer",
            )
        )

        // Protect sheet
        sheet.protect("invoice123")

        // Add fake logo image
        val fakeImage = ByteArray(100) { it.toByte() }
        sheet.addPicture(fakeImage, ImageFormat.PNG, 4, 0, 6, 2, 120, 60)

        // Column widths
        sheet.setColumnWidth(0, 20.0)
        sheet.setColumnWidth(1, 30.0)
        sheet.setColumnWidth(2, 8.0)
        sheet.setColumnWidth(3, 12.0)
        sheet.setColumnWidth(4, 14.0)

        // === Recalculate ===
        val evaluator = FormulaEvaluator(sheet)
        evaluator.recalculateAll()

        // === Save and reopen ===
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        val s = wb2.sheets[0]

        // Verify
        assertEquals("INVOICE", s["A1"].stringValue)
        assertEquals("Acme Corporation", s["A4"].stringValue)
        assertEquals("DroidXLS License", s.cell(9, 0).stringValue)
        assertEquals(1, s.mergedRegions.size)
        assertEquals("A1:F1", s.mergedRegions[0].toA1())

        // Verify formula preserved
        assertNotNull(s.cell(9, 4).formula)

        // CSV export
        val csv = CsvConverter.convertToString(s)
        assertTrue(csv.contains("INVOICE"))
        assertTrue(csv.contains("DroidXLS License"))
    }

    @Test
    fun `password protected workflow - save encrypted, open with correct password, verify data`() {
        // Create
        val wb = Workbook()
        val sheet = wb.addSheet("Confidential")
        sheet["A1"].value = "Employee"
        sheet["B1"].value = "Salary"
        sheet["A2"].value = "Alice"
        sheet["B2"].value = 85000.0
        sheet["A3"].value = "Bob"
        sheet["B3"].value = 92000.0
        sheet["B4"].formula = "=AVERAGE(B2:B3)"

        val evaluator = FormulaEvaluator(sheet)
        evaluator.recalculateAll()

        // Save with password
        val buffer = ByteArrayOutputStream()
        wb.save(buffer, "payroll2025")

        // Reopen with password
        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()), "payroll2025")
        assertEquals("Confidential", wb2.sheets[0].name)
        assertEquals("Alice", wb2.sheets[0]["A2"].stringValue)
        assertEquals(85000.0, wb2.sheets[0]["B2"].numericValue)
        assertNotNull(wb2.sheets[0]["B4"].formula)

        // Export to CSV (from decrypted data)
        val csv = CsvConverter.convertToString(wb2.sheets[0])
        assertTrue(csv.contains("Alice"))
        assertTrue(csv.contains("85000"))
    }

    @Test
    fun `japanese content workflow - CJK text, styles, formulas, round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("売上レポート")

        sheet["A1"].value = "商品名"
        sheet["B1"].value = "単価"
        sheet["C1"].value = "数量"
        sheet["D1"].value = "小計"

        for (col in 0..3) {
            sheet.cell(0, col).style {
                font { bold = true; name = "Yu Gothic" }
                fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.LIGHT_GRAY }
            }
        }

        val items = listOf(
            Triple("ウィジェットA", 1500.0, 10.0),
            Triple("ウィジェットB", 2800.0, 5.0),
            Triple("サービスパック", 50000.0, 1.0),
        )

        for ((i, item) in items.withIndex()) {
            val row = i + 1
            sheet.cell(row, 0).value = item.first
            sheet.cell(row, 1).value = item.second
            sheet.cell(row, 2).value = item.third
            sheet.cell(row, 3).formula = "=B${row + 1}*C${row + 1}"
        }

        val totalRow = items.size + 1
        sheet.cell(totalRow, 2).value = "合計"
        sheet.cell(totalRow, 2).style { font { bold = true }; alignment { horizontal = HorizontalAlignment.RIGHT } }
        sheet.cell(totalRow, 3).formula = "=SUM(D2:D${totalRow})"

        // Recalculate
        val evaluator = FormulaEvaluator(sheet)
        evaluator.recalculateAll()

        // Verify subtotals
        val subtotal1 = (sheet.cell(1, 3).cellValue as CellValue.Formula).cachedValue as CellValue.Number
        assertEquals(15000.0, subtotal1.value) // 1500 * 10

        // Save and reopen
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))

        assertEquals("売上レポート", wb2.sheets[0].name)
        assertEquals("商品名", wb2.sheets[0]["A1"].stringValue)
        assertEquals("ウィジェットA", wb2.sheets[0]["A2"].stringValue)
        assertEquals(1500.0, wb2.sheets[0]["B2"].numericValue)

        // CSV export with Japanese
        val csv = CsvConverter.convertToString(wb2.sheets[0])
        assertTrue(csv.contains("商品名"))
        assertTrue(csv.contains("ウィジェットA"))

        // HTML export
        val html = HtmlConverter.convert(wb2.sheets[0], title = "売上レポート")
        assertTrue(html.contains("売上レポート"))
        assertTrue(html.contains("ウィジェットA"))
    }
}
