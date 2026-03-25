package com.droidoffice.xls.e2e

import com.droidoffice.core.drawingml.BorderStyle
import com.droidoffice.core.drawingml.OfficeColor
import com.droidoffice.core.drawingml.PatternType
import com.droidoffice.core.ooxml.OoxmlPackage
import com.droidoffice.xls.core.CellRange
import com.droidoffice.xls.core.Workbook
import com.droidoffice.xls.drawing.ImageFormat
import com.droidoffice.xls.format.NumberFormat
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import kotlin.test.assertEquals
import kotlin.test.assertNotNull
import kotlin.test.assertTrue

/**
 * Validates that generated .xlsx files have correct OOXML structure
 * (would be readable by Excel, Google Sheets, LibreOffice).
 */
class XlsxStructureTest {

    private fun saveAndInspect(wb: Workbook): OoxmlPackage {
        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        return OoxmlPackage.open(ByteArrayInputStream(buffer.toByteArray()))
    }

    @Test
    fun `minimal workbook has all required OOXML parts`() {
        val wb = Workbook()
        wb.addSheet("Sheet1")["A1"].value = "test"

        val pkg = saveAndInspect(wb)

        // Required parts per OOXML spec
        assertNotNull(pkg.getPart("[Content_Types].xml"), "Missing [Content_Types].xml")
        assertNotNull(pkg.getPart("_rels/.rels"), "Missing _rels/.rels")
        assertNotNull(pkg.getPart("xl/workbook.xml"), "Missing xl/workbook.xml")
        assertNotNull(pkg.getPart("xl/_rels/workbook.xml.rels"), "Missing workbook.xml.rels")
        assertNotNull(pkg.getPart("xl/styles.xml"), "Missing xl/styles.xml")
        assertNotNull(pkg.getPart("xl/worksheets/sheet1.xml"), "Missing sheet1.xml")
        assertNotNull(pkg.getPart("xl/sharedStrings.xml"), "Missing sharedStrings.xml")
    }

    @Test
    fun `Content_Types xml contains correct content types`() {
        val wb = Workbook()
        wb.addSheet("Sheet1")["A1"].value = "test"

        val pkg = saveAndInspect(wb)
        val ct = pkg.getPart("[Content_Types].xml")!!.decodeToString()

        assertTrue(ct.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"))
        assertTrue(ct.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))
        assertTrue(ct.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"))
        assertTrue(ct.contains("application/vnd.openxmlformats-package.relationships+xml"))
    }

    @Test
    fun `workbook xml contains sheet references`() {
        val wb = Workbook()
        wb.addSheet("Data")["A1"].value = "a"
        wb.addSheet("Summary")["A1"].value = "b"

        val pkg = saveAndInspect(wb)
        val wbXml = pkg.getPart("xl/workbook.xml")!!.decodeToString()

        assertTrue(wbXml.contains("name=\"Data\""))
        assertTrue(wbXml.contains("name=\"Summary\""))
        assertTrue(wbXml.contains("r:id=\"rId1\""))
        assertTrue(wbXml.contains("r:id=\"rId2\""))
    }

    @Test
    fun `relationships point to correct targets`() {
        val wb = Workbook()
        wb.addSheet("Sheet1")["A1"].value = "test"

        val pkg = saveAndInspect(wb)
        val rels = pkg.getPart("xl/_rels/workbook.xml.rels")!!.decodeToString()

        assertTrue(rels.contains("Target=\"worksheets/sheet1.xml\""))
        assertTrue(rels.contains("Target=\"styles.xml\""))
        assertTrue(rels.contains("Target=\"sharedStrings.xml\""))
    }

    @Test
    fun `shared strings xml contains all text values`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Hello"
        sheet["A2"].value = "World"
        sheet["A3"].value = "日本語テスト"

        val pkg = saveAndInspect(wb)
        val sst = pkg.getPart("xl/sharedStrings.xml")!!.decodeToString()

        assertTrue(sst.contains("<t>Hello</t>"))
        assertTrue(sst.contains("<t>World</t>"))
        assertTrue(sst.contains("<t>日本語テスト</t>"))
        assertTrue(sst.contains("count=\"3\""))
    }

    @Test
    fun `styles xml has valid structure`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Styled"
        sheet["A1"].style {
            font { bold = true; size = 14.0 }
            fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.RED }
            border { all = BorderStyle.THIN }
            numberFormat(NumberFormat.THOUSANDS_DECIMAL)
        }

        val pkg = saveAndInspect(wb)
        val styles = pkg.getPart("xl/styles.xml")!!.decodeToString()

        assertTrue(styles.contains("<fonts"))
        assertTrue(styles.contains("<fills"))
        assertTrue(styles.contains("<borders"))
        assertTrue(styles.contains("<cellXfs"))
        assertTrue(styles.contains("<b/>"), "Bold flag missing")
        assertTrue(styles.contains("val=\"14.0\""), "Font size missing")
    }

    @Test
    fun `sheet xml has valid cell references`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "text"
        sheet["B2"].value = 42.0
        sheet["C3"].value = true

        val pkg = saveAndInspect(wb)
        val sheetXml = pkg.getPart("xl/worksheets/sheet1.xml")!!.decodeToString()

        assertTrue(sheetXml.contains("r=\"A1\""))
        assertTrue(sheetXml.contains("r=\"B2\""))
        assertTrue(sheetXml.contains("r=\"C3\""))
        assertTrue(sheetXml.contains("t=\"s\""), "Text cell should have type 's'")
        assertTrue(sheetXml.contains("t=\"b\""), "Boolean cell should have type 'b'")
    }

    @Test
    fun `workbook with images has drawing parts`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Images")
        sheet["A1"].value = "Photo"
        val fakeImage = ByteArray(50) { it.toByte() }
        sheet.addPicture(fakeImage, ImageFormat.PNG, 1, 1, 4, 8, 200, 150)

        val pkg = saveAndInspect(wb)

        assertNotNull(pkg.getPart("xl/drawings/drawing1.xml"), "Missing drawing XML")
        assertNotNull(pkg.getPart("xl/drawings/_rels/drawing1.xml.rels"), "Missing drawing rels")

        val ct = pkg.getPart("[Content_Types].xml")!!.decodeToString()
        assertTrue(ct.contains("image/png"), "Content types should include PNG")
        assertTrue(ct.contains("drawing+xml"), "Content types should include drawing")

        // Check drawing XML references image
        val drawingXml = pkg.getPart("xl/drawings/drawing1.xml")!!.decodeToString()
        assertTrue(drawingXml.contains("r:embed="))
        assertTrue(drawingXml.contains("twoCellAnchor"))
    }

    @Test
    fun `hidden sheet has state attribute in workbook xml`() {
        val wb = Workbook()
        wb.addSheet("Visible")["A1"].value = "show"
        val hidden = wb.addSheet("Hidden")
        hidden["A1"].value = "hide"
        hidden.isHidden = true

        val pkg = saveAndInspect(wb)
        val wbXml = pkg.getPart("xl/workbook.xml")!!.decodeToString()

        assertTrue(wbXml.contains("state=\"hidden\""), "Hidden sheet should have state attribute")
    }

    @Test
    fun `xml special characters are properly escaped`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = "Tom & Jerry <friends> \"best\" 'pals'"
        sheet["A2"].value = "<script>alert('xss')</script>"

        val pkg = saveAndInspect(wb)
        val sst = pkg.getPart("xl/sharedStrings.xml")!!.decodeToString()

        assertTrue(sst.contains("&amp;"), "& should be escaped")
        assertTrue(sst.contains("&lt;"), "< should be escaped")
        assertTrue(sst.contains("&gt;"), "> should be escaped")
        assertTrue(!sst.contains("<script>"), "Raw HTML tags should not appear")
    }

    @Test
    fun `large workbook 1000 rows generates valid xlsx`() {
        val wb = Workbook()
        val sheet = wb.addSheet("LargeData")

        for (row in 0 until 1000) {
            sheet.cell(row, 0).value = "Row $row"
            sheet.cell(row, 1).value = row.toDouble()
            sheet.cell(row, 2).value = row * 1.5
        }

        val buffer = ByteArrayOutputStream()
        wb.save(buffer)
        assertTrue(buffer.size() > 0)

        // Reopen and verify
        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        assertEquals(1, wb2.sheetCount)
        assertEquals("Row 0", wb2.sheets[0]["A1"].stringValue)
        assertEquals("Row 999", wb2.sheets[0].cell(999, 0).stringValue)
        assertEquals(999.0, wb2.sheets[0].cell(999, 1).numericValue)
    }
}
