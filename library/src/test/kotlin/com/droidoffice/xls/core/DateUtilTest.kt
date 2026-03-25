package com.droidoffice.xls.core

import org.junit.jupiter.api.Test
import java.time.LocalDate
import java.time.LocalDateTime
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class DateUtilTest {

    @Test
    fun `serial 1 is 1900-01-01`() {
        val dt = DateUtil.serialToDateTime(1.0)
        assertEquals(LocalDate.of(1900, 1, 1), dt.toLocalDate())
    }

    @Test
    fun `serial 44927 is 2023-01-01`() {
        // Known value: Jan 1, 2023 = serial 44927
        val dt = DateUtil.serialToDateTime(44927.0)
        assertEquals(LocalDate.of(2023, 1, 1), dt.toLocalDate())
    }

    @Test
    fun `round trip date conversion`() {
        val original = LocalDate.of(2025, 6, 15)
        val serial = DateUtil.dateToSerial(original)
        val restored = DateUtil.serialToDateTime(serial).toLocalDate()
        assertEquals(original, restored)
    }

    @Test
    fun `round trip datetime with time portion`() {
        val original = LocalDateTime.of(2025, 3, 25, 14, 30, 0)
        val serial = DateUtil.dateTimeToSerial(original)
        val restored = DateUtil.serialToDateTime(serial)
        assertEquals(original.toLocalDate(), restored.toLocalDate())
        assertEquals(original.hour, restored.hour)
        assertEquals(original.minute, restored.minute)
    }

    @Test
    fun `built-in date formats detected`() {
        assertTrue(DateUtil.isDateFormat(14))  // m/d/yyyy
        assertTrue(DateUtil.isDateFormat(22))  // m/d/yyyy h:mm
    }

    @Test
    fun `custom date format string detected`() {
        assertTrue(DateUtil.isDateFormatString("yyyy-mm-dd"))
        assertTrue(DateUtil.isDateFormatString("dd/mm/yyyy"))
        assertTrue(DateUtil.isDateFormatString("h:mm:ss"))
    }
}
