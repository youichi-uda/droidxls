package com.droidoffice.xls.core

import org.junit.jupiter.api.Test
import kotlin.test.assertEquals

class WorkbookTest {

    @Test
    fun `move sheet changes order`() {
        val wb = Workbook()
        wb.addSheet("A")
        wb.addSheet("B")
        wb.addSheet("C")

        wb.moveSheet(2, 0) // Move C to front

        assertEquals("C", wb.sheets[0].name)
        assertEquals("A", wb.sheets[1].name)
        assertEquals("B", wb.sheets[2].name)
    }

    @Test
    fun `copy sheet duplicates data`() {
        val wb = Workbook()
        val original = wb.addSheet("Original")
        original["A1"].value = "Hello"
        original["B1"].value = 42.0
        original.setColumnWidth(0, 15.0)
        original.hideRow(3)

        val copy = wb.copySheet(0, "Copy")

        assertEquals("Hello", copy["A1"].stringValue)
        assertEquals(42.0, copy["B1"].numericValue)
        assertEquals(15.0, copy.columnWidths[0])
        assertEquals(true, copy.isRowHidden(3))
        assertEquals(2, wb.sheetCount)
    }

    @Test
    fun `active sheet index`() {
        val wb = Workbook()
        wb.addSheet("A")
        wb.addSheet("B")
        wb.addSheet("C")

        assertEquals(0, wb.activeSheetIndex)

        wb.activeSheetIndex = 2
        assertEquals(2, wb.activeSheetIndex)
        assertEquals("C", wb.activeSheet?.name)
    }

    @Test
    fun `insert sheet at position`() {
        val wb = Workbook()
        wb.addSheet("A")
        wb.addSheet("C")
        wb.insertSheet(1, "B")

        assertEquals("A", wb.sheets[0].name)
        assertEquals("B", wb.sheets[1].name)
        assertEquals("C", wb.sheets[2].name)
    }
}
