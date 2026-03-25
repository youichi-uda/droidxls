package com.droidoffice.xls.format

import com.droidoffice.xls.core.DateUtil

/**
 * Represents the workbook's style sheet (styles.xml).
 * Holds enough information to determine cell formatting (especially date detection).
 */
class StyleSheet {
    /** Custom number formats: numFmtId → format string */
    val numberFormats = mutableMapOf<Int, String>()

    /** Cell XF entries (index matches the 's' attribute on <c>). Each entry holds a numFmtId. */
    val cellXfs = mutableListOf<CellXf>()

    /**
     * Determine if a cell style index represents a date format.
     */
    fun isDateFormat(styleIndex: Int): Boolean {
        if (styleIndex !in cellXfs.indices) return false
        val xf = cellXfs[styleIndex]
        val numFmtId = xf.numFmtId

        // Check built-in date format IDs
        if (DateUtil.isDateFormat(numFmtId)) return true

        // Check custom format string
        val formatString = numberFormats[numFmtId] ?: return false
        return DateUtil.isDateFormatString(formatString)
    }
}

/**
 * Minimal cell XF (formatting) record.
 */
data class CellXf(
    val numFmtId: Int = 0,
    val fontId: Int = 0,
    val fillId: Int = 0,
    val borderId: Int = 0,
)
