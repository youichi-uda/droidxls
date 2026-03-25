package com.droidoffice.xls.core

import org.junit.jupiter.api.Test
import kotlin.test.assertEquals
import kotlin.test.assertNull
import kotlin.test.assertTrue

class WorksheetTest {

    @Test
    fun `insert rows shifts cells down`() {
        val sheet = Workbook().addSheet("Test")
        sheet["A1"].value = "Header"
        sheet["A2"].value = "Data1"
        sheet["A3"].value = "Data2"

        sheet.insertRows(1, 2) // Insert 2 rows at row index 1

        assertEquals("Header", sheet["A1"].stringValue)
        assertEquals("Data1", sheet["A4"].stringValue)
        assertEquals("Data2", sheet["A5"].stringValue)
    }

    @Test
    fun `delete rows shifts cells up`() {
        val sheet = Workbook().addSheet("Test")
        sheet["A1"].value = "Keep"
        sheet["A2"].value = "Delete"
        sheet["A3"].value = "Also Delete"
        sheet["A4"].value = "Move Up"

        sheet.deleteRows(1, 2) // Delete rows at index 1 and 2

        assertEquals("Keep", sheet["A1"].stringValue)
        assertEquals("Move Up", sheet["A2"].stringValue)
    }

    @Test
    fun `insert columns shifts cells right`() {
        val sheet = Workbook().addSheet("Test")
        sheet["A1"].value = "A"
        sheet["B1"].value = "B"
        sheet["C1"].value = "C"

        sheet.insertColumns(1) // Insert 1 column at B

        assertEquals("A", sheet["A1"].stringValue)
        assertEquals("B", sheet["C1"].stringValue)
        assertEquals("C", sheet["D1"].stringValue)
    }

    @Test
    fun `delete columns shifts cells left`() {
        val sheet = Workbook().addSheet("Test")
        sheet["A1"].value = "A"
        sheet["B1"].value = "Delete"
        sheet["C1"].value = "C"

        sheet.deleteColumns(1) // Delete column B

        assertEquals("A", sheet["A1"].stringValue)
        assertEquals("C", sheet["B1"].stringValue)
    }

    @Test
    fun `hide and show rows`() {
        val sheet = Workbook().addSheet("Test")
        sheet.hideRow(2)
        assertTrue(sheet.isRowHidden(2))
        sheet.showRow(2)
        assertTrue(!sheet.isRowHidden(2))
    }

    @Test
    fun `hide and show columns`() {
        val sheet = Workbook().addSheet("Test")
        sheet.hideColumn(1)
        assertTrue(sheet.isColumnHidden(1))
        sheet.showColumn(1)
        assertTrue(!sheet.isColumnHidden(1))
    }

    @Test
    fun `freeze and unfreeze`() {
        val sheet = Workbook().addSheet("Test")
        sheet.freeze(1, 2)
        assertEquals(CellReference(1, 2), sheet.freezePane)
        sheet.unfreeze()
        assertNull(sheet.freezePane)
    }
}
