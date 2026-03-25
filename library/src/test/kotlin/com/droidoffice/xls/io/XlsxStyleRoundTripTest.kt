package com.droidoffice.xls.io

import com.droidoffice.core.drawingml.BorderStyle
import com.droidoffice.core.drawingml.OfficeColor
import com.droidoffice.core.drawingml.PatternType
import com.droidoffice.xls.core.Workbook
import com.droidoffice.xls.format.HorizontalAlignment
import com.droidoffice.xls.format.NumberFormat
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class XlsxStyleRoundTripTest {

    private fun roundTrip(wb: Workbook): Workbook {
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        return Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
    }

    @Test
    fun `bold font style produces non-default style index`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Normal"
        sheet["B1"].value = "Bold"
        sheet["B1"].style {
            font { bold = true }
        }

        // After save, bold cell should have styleIndex > 0
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        val normalStyle = wb2.sheets[0]["A1"].styleIndex
        val boldStyle = wb2.sheets[0]["B1"].styleIndex
        assertTrue(boldStyle > normalStyle, "Bold cell should have higher style index than normal")
    }

    @Test
    fun `styled workbook can be opened by reader without errors`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Styled")
        sheet["A1"].value = "Header"
        sheet["A1"].style {
            font { bold = true; size = 14.0; color = OfficeColor.BLUE }
            fill {
                patternType = PatternType.SOLID
                foregroundColor = OfficeColor.LIGHT_BLUE
            }
            border { all = BorderStyle.THIN; allColor = OfficeColor.BLACK }
            alignment { horizontal = HorizontalAlignment.CENTER; wrapText = true }
        }

        sheet["A2"].value = 12345.67
        sheet["A2"].style {
            numberFormat(NumberFormat.THOUSANDS_DECIMAL)
        }

        // Round trip should not throw
        val wb2 = roundTrip(wb)
        assertEquals("Header", wb2.sheets[0]["A1"].stringValue)
        assertEquals(12345.67, wb2.sheets[0]["A2"].numericValue)
    }

    @Test
    fun `multiple cells with same style share style index`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        for (row in 0..9) {
            sheet.cell(row, 0).value = "Row $row"
            sheet.cell(row, 0).style {
                font { bold = true }
            }
        }

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        val indices = (0..9).map { wb2.sheets[0].cell(it, 0).styleIndex }.toSet()
        // All should share the same style index
        assertEquals(1, indices.size, "All cells with same style should share one index")
    }
}
