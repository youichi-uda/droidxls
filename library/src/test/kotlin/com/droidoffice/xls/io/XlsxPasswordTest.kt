package com.droidoffice.xls.io

import com.droidoffice.core.exception.PasswordException
import com.droidoffice.xls.core.Workbook
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.assertThrows
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import kotlin.test.assertEquals

class XlsxPasswordTest {

    @Test
    fun `save and open with password`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Secret")
        sheet["A1"].value = "Confidential"
        sheet["B1"].value = 42.0

        val buffer = ByteArrayOutputStream()
        wb.save(buffer, "mypassword")

        val wb2 = Workbook.open(ByteArrayInputStream(buffer.toByteArray()), "mypassword")
        assertEquals("Secret", wb2.sheets[0].name)
        assertEquals("Confidential", wb2.sheets[0]["A1"].stringValue)
        assertEquals(42.0, wb2.sheets[0]["B1"].numericValue)
    }

    @Test
    fun `open encrypted file without password throws`() {
        val wb = Workbook()
        wb.addSheet("Test")["A1"].value = "data"

        val buffer = ByteArrayOutputStream()
        wb.save(buffer, "pass123")

        assertThrows<PasswordException> {
            Workbook.open(ByteArrayInputStream(buffer.toByteArray()))
        }
    }

    @Test
    fun `wrong password throws PasswordException`() {
        val wb = Workbook()
        wb.addSheet("Test")["A1"].value = "data"

        val buffer = ByteArrayOutputStream()
        wb.save(buffer, "correct")

        assertThrows<PasswordException> {
            Workbook.open(ByteArrayInputStream(buffer.toByteArray()), "wrong")
        }
    }
}
