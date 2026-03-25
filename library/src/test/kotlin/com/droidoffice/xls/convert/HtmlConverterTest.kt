package com.droidoffice.xls.convert

import com.droidoffice.xls.core.Workbook
import org.junit.jupiter.api.Test
import kotlin.test.assertTrue

class HtmlConverterTest {

    @Test
    fun `basic HTML conversion`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Name"
        sheet["B1"].value = "Score"
        sheet["A2"].value = "Alice"
        sheet["B2"].value = 95.0

        val html = HtmlConverter.convert(sheet, title = "Test Sheet")
        assertTrue(html.contains("<title>Test Sheet</title>"))
        assertTrue(html.contains("<th>Name</th>"))
        assertTrue(html.contains("<th>Score</th>"))
        assertTrue(html.contains("<td>Alice</td>"))
        assertTrue(html.contains("<td>95</td>"))
    }

    @Test
    fun `HTML escapes special characters`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "<script>alert('xss')</script>"

        val html = HtmlConverter.convert(sheet)
        assertTrue(html.contains("&lt;script&gt;"))
        assertTrue(!html.contains("<script>"))
    }

    @Test
    fun `hidden rows and columns are skipped`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Visible"
        sheet["A2"].value = "Hidden"
        sheet["B1"].value = "Also Visible"
        sheet.hideRow(1)

        val html = HtmlConverter.convert(sheet)
        assertTrue(html.contains("Visible"))
        assertTrue(!html.contains("Hidden"))
    }
}
