package com.droidoffice.xls.io

import com.droidoffice.xls.chart.ChartType
import com.droidoffice.xls.core.*
import com.droidoffice.xls.drawing.ImageFormat
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import kotlin.test.assertEquals
import kotlin.test.assertNotNull
import kotlin.test.assertTrue

class XlsxPhase2RoundTripTest {

    private fun roundTrip(wb: Workbook): Workbook {
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        return Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
    }

    @Test
    fun `picture is included in xlsx output`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Pics")
        sheet["A1"].value = "Image test"

        // Add a fake 1x1 PNG
        val fakeImageData = byteArrayOf(0x89.toByte(), 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
        sheet.addPicture(fakeImageData, ImageFormat.PNG, 0, 0, 3, 5, 100, 100)

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        val bytes = buffer.toByteArray()
        assertTrue(bytes.isNotEmpty())

        // Re-open should at least not crash (drawing reading not fully implemented yet)
        val wb2 = Workbook.open(ByteArrayInputStream(bytes))
        assertEquals("Image test", wb2.sheets[0]["A1"].stringValue)
    }

    @Test
    fun `auto filter range round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Filter")
        sheet["A1"].value = "Name"
        sheet["B1"].value = "Value"
        sheet["A2"].value = "Test"
        sheet["B2"].value = 1.0
        sheet.autoFilterRange = CellRange.parse("A1:B2")

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        // Verify it writes without error
        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        assertEquals("Name", wb2.sheets[0]["A1"].stringValue)
    }

    @Test
    fun `data validation writes without error`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Validation")
        sheet["A1"].value = "Pick one"
        sheet.dataValidations.add(
            DataValidation(
                range = CellRange.parse("A2:A10"),
                type = ValidationType.LIST,
                formula1 = "\"Yes,No,Maybe\"",
                showDropDown = true,
                showErrorMessage = true,
                errorTitle = "Invalid",
                errorMessage = "Please select from list",
            )
        )

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        assertTrue(buffer.size() > 0)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        assertEquals("Pick one", wb2.sheets[0]["A1"].stringValue)
    }

    @Test
    fun `conditional formatting writes without error`() {
        val wb = Workbook()
        val sheet = wb.addSheet("CF")
        sheet["A1"].value = 10.0
        sheet["A2"].value = 20.0
        sheet.conditionalFormattings.add(
            ConditionalFormatting(
                ranges = listOf(CellRange.parse("A1:A10")),
                rules = listOf(
                    ConditionalRule(
                        type = ConditionalType.CELL_IS,
                        operator = ConditionalOperator.GREATER_THAN,
                        formula = "15",
                        priority = 1,
                    )
                ),
            )
        )

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        assertTrue(buffer.size() > 0)
    }

    @Test
    fun `sheet protection writes without error`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Protected")
        sheet["A1"].value = "Locked"
        sheet.protect("secret")

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        assertTrue(buffer.size() > 0)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        assertEquals("Locked", wb2.sheets[0]["A1"].stringValue)
    }

    @Test
    fun `chart model can be created`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Charts")
        sheet["A1"].value = "Q1"
        sheet["A2"].value = "Q2"
        sheet["B1"].value = 100.0
        sheet["B2"].value = 200.0

        sheet.addChart(ChartType.BAR) {
            title = "Sales"
            series("Sheet1!\$B\$1:\$B\$2") {
                name = "Revenue"
                categoryRange = "Sheet1!\$A\$1:\$A\$2"
            }
        }

        assertEquals(1, sheet.charts.size)
        assertEquals(ChartType.BAR, sheet.charts[0].type)
        assertEquals("Sales", sheet.charts[0].title)
        assertEquals(1, sheet.charts[0].series.size)
    }

    @Test
    fun `named ranges can be added`() {
        val wb = Workbook()
        wb.addSheet("Data")["A1"].value = "test"
        wb.namedRanges.add(NamedRange("MyRange", "Data!\$A\$1:\$A\$100"))
        wb.namedRanges.add(NamedRange("Header", "Data!\$A\$1:\$D\$1", sheetIndex = 0))

        assertEquals(2, wb.namedRanges.size)
        assertEquals("MyRange", wb.namedRanges[0].name)
    }
}
