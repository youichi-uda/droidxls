package com.droidoffice.xls.convert

import com.droidoffice.xls.core.Workbook
import org.junit.jupiter.api.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class CsvConverterTest {

    @Test
    fun `basic CSV conversion`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Name"
        sheet["B1"].value = "Score"
        sheet["A2"].value = "Alice"
        sheet["B2"].value = 95.5

        val csv = CsvConverter.convertToString(sheet)
        val lines = csv.trim().lines()
        assertEquals(2, lines.size)
        assertEquals("Name,Score", lines[0])
        assertEquals("Alice,95.5", lines[1])
    }

    @Test
    fun `CSV escapes commas and quotes`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Hello, World"
        sheet["B1"].value = "She said \"hi\""

        val csv = CsvConverter.convertToString(sheet)
        assertTrue(csv.contains("\"Hello, World\""))
        assertTrue(csv.contains("\"She said \"\"hi\"\"\""))
    }

    @Test
    fun `CSV with custom delimiter`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "A"
        sheet["B1"].value = "B"

        val csv = CsvConverter.convertToString(sheet, delimiter = '\t')
        assertTrue(csv.contains("A\tB"))
    }

    @Test
    fun `empty sheet produces empty CSV`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Empty")
        val csv = CsvConverter.convertToString(sheet)
        assertEquals("", csv)
    }
}
