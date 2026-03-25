package com.droidoffice.xls.core

import java.time.LocalDate
import java.time.LocalDateTime
import java.time.LocalTime
import java.time.temporal.ChronoUnit

/**
 * Converts between Excel serial date numbers and Java date/time types.
 *
 * Excel serial number 1 = January 1, 1900.
 * Due to a Lotus 1-2-3 bug, Excel treats 1900 as a leap year, so serial 60 = Feb 29, 1900
 * (which never existed). We handle this by adjusting for serials > 59.
 */
object DateUtil {

    // Day before serial 1 (Jan 1, 1900)
    private val EPOCH = LocalDate.of(1899, 12, 31)

    /**
     * Convert an Excel serial number to a LocalDateTime.
     */
    fun serialToDateTime(serial: Double): LocalDateTime {
        val wholeDays = serial.toLong()
        val fractionalDay = serial - wholeDays

        val date = when {
            wholeDays <= 0 -> EPOCH
            wholeDays <= 59 -> EPOCH.plusDays(wholeDays)
            wholeDays == 60L -> LocalDate.of(1900, 2, 28) // Map fictitious Feb 29 to Feb 28
            else -> EPOCH.plusDays(wholeDays - 1) // Subtract 1 to skip the fake Feb 29
        }

        val totalSeconds = Math.round(fractionalDay * 86400)
        val time = LocalTime.ofSecondOfDay(totalSeconds.coerceIn(0, 86399))

        return LocalDateTime.of(date, time)
    }

    /**
     * Convert a LocalDateTime to an Excel serial number.
     */
    fun dateTimeToSerial(dt: LocalDateTime): Double {
        var days = ChronoUnit.DAYS.between(EPOCH, dt.toLocalDate())
        // Add 1 for dates after Feb 28, 1900 to account for the fake Feb 29
        if (days > 59) days += 1

        val timeOfDay = dt.toLocalTime().toSecondOfDay().toDouble() / 86400.0
        return days.toDouble() + timeOfDay
    }

    /**
     * Convert a LocalDate to an Excel serial number (integer, no time portion).
     */
    fun dateToSerial(date: LocalDate): Double {
        return dateTimeToSerial(date.atStartOfDay())
    }

    /**
     * Heuristic: determine if a number format ID is a date format.
     * Built-in format IDs 14-22 and 45-47 are date/time formats.
     */
    fun isDateFormat(numFmtId: Int): Boolean {
        return numFmtId in 14..22 || numFmtId in 45..47
    }

    /**
     * Heuristic: determine if a custom format string looks like a date format.
     */
    fun isDateFormatString(formatString: String): Boolean {
        val lower = formatString.lowercase()
        // Remove quoted strings and color/condition blocks
        val stripped = lower.replace(Regex("\"[^\"]*\""), "").replace(Regex("\\[[^]]*]"), "")
        // Check for date/time tokens
        return stripped.contains("y") || (stripped.contains("m") && !stripped.contains("#")) ||
            stripped.contains("d") || stripped.contains("h") || stripped.contains("s")
    }
}
