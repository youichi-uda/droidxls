package com.droidoffice.xls.core

import org.junit.jupiter.api.Test
import kotlin.test.assertEquals

class CellReferenceTest {

    @Test
    fun `parse A1 reference`() {
        val ref = CellReference.parse("A1")
        assertEquals(0, ref.row)
        assertEquals(0, ref.col)
    }

    @Test
    fun `parse B2 reference`() {
        val ref = CellReference.parse("B2")
        assertEquals(1, ref.row)
        assertEquals(1, ref.col)
    }

    @Test
    fun `parse AA100 reference`() {
        val ref = CellReference.parse("AA100")
        assertEquals(99, ref.row)
        assertEquals(26, ref.col)
    }

    @Test
    fun `toA1 round trip`() {
        val ref = CellReference(0, 0)
        assertEquals("A1", ref.toA1())

        val ref2 = CellReference(99, 26)
        assertEquals("AA100", ref2.toA1())
    }

    @Test
    fun `column letters conversion`() {
        assertEquals("A", CellReference.columnToLetters(0))
        assertEquals("Z", CellReference.columnToLetters(25))
        assertEquals("AA", CellReference.columnToLetters(26))
        assertEquals("AZ", CellReference.columnToLetters(51))
        assertEquals("BA", CellReference.columnToLetters(52))
    }
}
