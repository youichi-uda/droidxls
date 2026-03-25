package com.droidoffice.xls.io

import com.droidoffice.xls.core.CellRange
import com.droidoffice.xls.core.CellReference
import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.Workbook
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class XlsxRoundTripTest {

    @Test
    fun `create workbook with text and numbers, write and read back`() {
        // Create
        val wb = Workbook()
        val sheet = wb.addSheet("Sheet1")
        sheet["A1"].value = "Name"
        sheet["B1"].value = "Score"
        sheet["A2"].value = "Alice"
        sheet["B2"].value = 95.5
        sheet["A3"].value = "Bob"
        sheet["B3"].value = 87.0

        // Write
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        val bytes = buffer.toByteArray()
        assertTrue(bytes.isNotEmpty())

        // Read back
        val wb2 = Workbook.open(ByteArrayInputStream(bytes))
        assertEquals(1, wb2.sheetCount)
        assertEquals("Sheet1", wb2.sheets[0].name)

        val s = wb2.sheets[0]
        assertEquals("Name", s["A1"].stringValue)
        assertEquals("Score", s["B1"].stringValue)
        assertEquals("Alice", s["A2"].stringValue)
        assertEquals(95.5, s["B2"].numericValue)
        assertEquals("Bob", s["A3"].stringValue)
        assertEquals(87.0, s["B3"].numericValue)
    }

    @Test
    fun `boolean values round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = true
        sheet["A2"].value = false

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        val s = wb2.sheets[0]
        assertTrue(s["A1"].cellValue is CellValue.Bool)
        assertEquals(true, (s["A1"].cellValue as CellValue.Bool).value)
        assertEquals(false, (s["A2"].cellValue as CellValue.Bool).value)
    }

    @Test
    fun `formula values round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = 10.0
        sheet["A2"].value = 20.0
        sheet["A3"].formula = "=SUM(A1:A2)"

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        val s = wb2.sheets[0]
        assertEquals("SUM(A1:A2)", s["A3"].formula)
    }

    @Test
    fun `multiple sheets round trip`() {
        val wb = Workbook()
        wb.addSheet("Sheet1")["A1"].value = "One"
        wb.addSheet("Sheet2")["A1"].value = "Two"
        wb.addSheet("Data")["A1"].value = "Three"

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        assertEquals(3, wb2.sheetCount)
        assertEquals("Sheet1", wb2.sheets[0].name)
        assertEquals("Sheet2", wb2.sheets[1].name)
        assertEquals("Data", wb2.sheets[2].name)
        assertEquals("One", wb2.sheets[0]["A1"].stringValue)
        assertEquals("Two", wb2.sheets[1]["A1"].stringValue)
        assertEquals("Three", wb2.sheets[2]["A1"].stringValue)
    }

    @Test
    fun `merged cells round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Merged"
        sheet.addMergedRegion(CellRange.parse("A1:C1"))

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        val s = wb2.sheets[0]
        assertEquals(1, s.mergedRegions.size)
        assertEquals("A1:C1", s.mergedRegions[0].toA1())
    }

    @Test
    fun `japanese text round trip`() {
        val wb = Workbook()
        val sheet = wb.addSheet("日本語テスト")
        sheet["A1"].value = "名前"
        sheet["B1"].value = "売上"
        sheet["A2"].value = "田中太郎"
        sheet["B2"].value = 1234567.89

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        assertEquals("日本語テスト", wb2.sheets[0].name)
        assertEquals("名前", wb2.sheets[0]["A1"].stringValue)
        assertEquals("田中太郎", wb2.sheets[0]["A2"].stringValue)
    }
}
