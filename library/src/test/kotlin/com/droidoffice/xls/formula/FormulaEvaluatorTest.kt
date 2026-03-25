package com.droidoffice.xls.formula

import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.ErrorCode
import com.droidoffice.xls.core.Workbook
import org.junit.jupiter.api.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FormulaEvaluatorTest {

    private fun evalFormula(formula: String, setup: (com.droidoffice.xls.core.Worksheet) -> Unit = {}): CellValue {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        setup(sheet)
        sheet["Z1"].formula = formula
        val evaluator = FormulaEvaluator(sheet)
        return evaluator.evaluate(sheet["Z1"])
    }

    // -- Aggregate --
    @Test fun `SUM of range`() {
        val result = evalFormula("SUM(A1:A3)") { s ->
            s["A1"].value = 10.0; s["A2"].value = 20.0; s["A3"].value = 30.0
        }
        assertEquals(60.0, (result as CellValue.Number).value)
    }

    @Test fun `AVERAGE`() {
        val result = evalFormula("AVERAGE(A1:A3)") { s ->
            s["A1"].value = 10.0; s["A2"].value = 20.0; s["A3"].value = 30.0
        }
        assertEquals(20.0, (result as CellValue.Number).value)
    }

    @Test fun `COUNT and COUNTA`() {
        val r1 = evalFormula("COUNT(A1:A3)") { s ->
            s["A1"].value = 10.0; s["A2"].value = "text"; s["A3"].value = 30.0
        }
        assertEquals(2.0, (r1 as CellValue.Number).value)

        val r2 = evalFormula("COUNTA(A1:A3)") { s ->
            s["A1"].value = 10.0; s["A2"].value = "text"; s["A3"].value = 30.0
        }
        assertEquals(3.0, (r2 as CellValue.Number).value)
    }

    @Test fun `MAX and MIN`() {
        val max = evalFormula("MAX(A1:A3)") { s ->
            s["A1"].value = 5.0; s["A2"].value = 15.0; s["A3"].value = 10.0
        }
        assertEquals(15.0, (max as CellValue.Number).value)

        val min = evalFormula("MIN(A1:A3)") { s ->
            s["A1"].value = 5.0; s["A2"].value = 15.0; s["A3"].value = 10.0
        }
        assertEquals(5.0, (min as CellValue.Number).value)
    }

    // -- Logic --
    @Test fun `IF true and false`() {
        val t = evalFormula("IF(A1>5,\"big\",\"small\")") { it["A1"].value = 10.0 }
        assertEquals("big", (t as CellValue.Text).value)

        val f = evalFormula("IF(A1>5,\"big\",\"small\")") { it["A1"].value = 3.0 }
        assertEquals("small", (f as CellValue.Text).value)
    }

    @Test fun `AND and OR`() {
        val and = evalFormula("AND(A1>0,A2>0)") { it["A1"].value = 1.0; it["A2"].value = -1.0 }
        assertEquals(false, (and as CellValue.Bool).value)

        val or = evalFormula("OR(A1>0,A2>0)") { it["A1"].value = 1.0; it["A2"].value = -1.0 }
        assertEquals(true, (or as CellValue.Bool).value)
    }

    @Test fun `IFERROR`() {
        val r = evalFormula("IFERROR(A1/A2,0)") { it["A1"].value = 10.0; it["A2"].value = 0.0 }
        // A1/A2 produces DIV/0 error, IFERROR should return 0
        assertEquals(0.0, (r as CellValue.Number).value)
    }

    // -- Math --
    @Test fun `ROUND`() {
        val r = evalFormula("ROUND(3.14159,2)") {}
        assertEquals(3.14, (r as CellValue.Number).value)
    }

    @Test fun `ABS`() {
        val r = evalFormula("ABS(-42)") {}
        assertEquals(42.0, (r as CellValue.Number).value)
    }

    @Test fun `MOD`() {
        val r = evalFormula("MOD(10,3)") {}
        assertEquals(1.0, (r as CellValue.Number).value)
    }

    @Test fun `INT`() {
        val r = evalFormula("INT(7.8)") {}
        assertEquals(7.0, (r as CellValue.Number).value)
    }

    // -- String --
    @Test fun `CONCATENATE`() {
        val r = evalFormula("CONCATENATE(A1,\" \",A2)") { it["A1"].value = "Hello"; it["A2"].value = "World" }
        assertEquals("Hello World", (r as CellValue.Text).value)
    }

    @Test fun `LEFT RIGHT MID`() {
        val l = evalFormula("LEFT(\"Hello\",3)") {}
        assertEquals("Hel", (l as CellValue.Text).value)

        val r = evalFormula("RIGHT(\"Hello\",2)") {}
        assertEquals("lo", (r as CellValue.Text).value)

        val m = evalFormula("MID(\"Hello\",2,3)") {}
        assertEquals("ell", (m as CellValue.Text).value)
    }

    @Test fun `LEN TRIM UPPER LOWER`() {
        assertEquals(5.0, (evalFormula("LEN(\"Hello\")") {} as CellValue.Number).value)
        assertEquals("Hello", (evalFormula("TRIM(\"  Hello  \")") {} as CellValue.Text).value)
        assertEquals("HELLO", (evalFormula("UPPER(\"Hello\")") {} as CellValue.Text).value)
        assertEquals("hello", (evalFormula("LOWER(\"Hello\")") {} as CellValue.Text).value)
    }

    @Test fun `SUBSTITUTE`() {
        val r = evalFormula("SUBSTITUTE(\"Hello World\",\"World\",\"Kotlin\")") {}
        assertEquals("Hello Kotlin", (r as CellValue.Text).value)
    }

    // -- Date --
    @Test fun `DATE function`() {
        val r = evalFormula("DATE(2025,3,25)") {}
        assertTrue(r is CellValue.DateValue)
        assertEquals(2025, (r as CellValue.DateValue).value.year)
        assertEquals(3, r.value.monthValue)
        assertEquals(25, r.value.dayOfMonth)
    }

    @Test fun `YEAR MONTH DAY`() {
        val setup: (com.droidoffice.xls.core.Worksheet) -> Unit = {
            it["A1"].cellValue = CellValue.DateValue(java.time.LocalDate.of(2025, 6, 15).atStartOfDay())
        }
        assertEquals(2025.0, (evalFormula("YEAR(A1)", setup) as CellValue.Number).value)
        assertEquals(6.0, (evalFormula("MONTH(A1)", setup) as CellValue.Number).value)
        assertEquals(15.0, (evalFormula("DAY(A1)", setup) as CellValue.Number).value)
    }

    // -- Arithmetic --
    @Test fun `basic arithmetic`() {
        val r = evalFormula("(A1+A2)*2") { it["A1"].value = 5.0; it["A2"].value = 3.0 }
        assertEquals(16.0, (r as CellValue.Number).value)
    }

    @Test fun `string concatenation with ampersand`() {
        val r = evalFormula("A1&\" \"&A2") { it["A1"].value = "Hello"; it["A2"].value = "World" }
        assertEquals("Hello World", (r as CellValue.Text).value)
    }

    @Test fun `division by zero`() {
        val r = evalFormula("A1/0") { it["A1"].value = 10.0 }
        assertTrue(r is CellValue.Error)
        assertEquals(ErrorCode.DIV_ZERO, (r as CellValue.Error).code)
    }

    @Test fun `unknown function returns NAME error`() {
        val r = evalFormula("NOSUCHFUNC(1)") {}
        assertTrue(r is CellValue.Error)
        assertEquals(ErrorCode.NAME, (r as CellValue.Error).code)
    }

    @Test fun `recalculateAll updates cached values`() {
        val wb = Workbook()
        val sheet = wb.addSheet("Test")
        sheet["A1"].value = 10.0
        sheet["A2"].value = 20.0
        sheet["A3"].formula = "=SUM(A1:A2)"

        val evaluator = FormulaEvaluator(sheet)
        evaluator.recalculateAll()

        val formula = sheet["A3"].cellValue as CellValue.Formula
        assertEquals(30.0, (formula.cachedValue as CellValue.Number).value)
    }
}
