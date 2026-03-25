package com.droidoffice.xls.io

import com.droidoffice.xls.core.CellReference
import com.droidoffice.xls.core.Workbook
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class XlsxPhase1bRoundTripTest {

    private fun roundTrip(wb: Workbook): Workbook {
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        return Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
    }

    @Test
    fun `column widths round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "data"
        sheet.setColumnWidth(0, 20.0)
        sheet.setColumnWidth(2, 35.5)

        val wb2 = roundTrip(wb)
        val s = wb2.sheets[0]
        assertEquals(20.0, s.columnWidths[0])
        assertEquals(35.5, s.columnWidths[2])
    }

    @Test
    fun `row heights round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "data"
        sheet.setRowHeight(0, 30.0)

        val wb2 = roundTrip(wb)
        val s = wb2.sheets[0]
        assertEquals(30.0, s.rowHeights[0])
    }

    @Test
    fun `hidden rows round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "visible"
        sheet["A2"].value = "hidden"
        sheet.hideRow(1)

        val wb2 = roundTrip(wb)
        assertTrue(wb2.sheets[0].isRowHidden(1))
    }

    @Test
    fun `hidden columns round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "visible"
        sheet["B1"].value = "hidden"
        sheet.hideColumn(1)

        val wb2 = roundTrip(wb)
        assertTrue(wb2.sheets[0].isColumnHidden(1))
    }

    @Test
    fun `freeze pane round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "frozen"
        sheet.freeze(1, 2)

        val wb2 = roundTrip(wb)
        val fp = wb2.sheets[0].freezePane
        assertEquals(CellReference(1, 2), fp)
    }

    @Test
    fun `hidden sheet round trip`() {
        val wb = Workbook()
        wb.addSheet("Visible")["A1"].value = "see me"
        val hidden = wb.addSheet("Hidden")
        hidden["A1"].value = "can't see"
        hidden.isHidden = true

        val wb2 = roundTrip(wb)
        assertEquals(false, wb2.sheets[0].isHidden)
        assertEquals(true, wb2.sheets[1].isHidden)
    }
}
