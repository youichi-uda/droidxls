package com.droidoffice.xls.format

/**
 * Number format for a cell.
 */
data class NumberFormat(
    val id: Int = 0,
    val formatCode: String = "General",
) {
    companion object {
        val GENERAL = NumberFormat(0, "General")
        val INTEGER = NumberFormat(1, "0")
        val DECIMAL_2 = NumberFormat(2, "0.00")
        val THOUSANDS = NumberFormat(3, "#,##0")
        val THOUSANDS_DECIMAL = NumberFormat(4, "#,##0.00")
        val PERCENT = NumberFormat(9, "0%")
        val PERCENT_DECIMAL = NumberFormat(10, "0.00%")
        val DATE_MDY = NumberFormat(14, "m/d/yyyy")
        val DATE_DMY = NumberFormat(15, "d-mmm-yy")
        val DATE_FULL = NumberFormat(16, "d-mmm")
        val TIME_HM = NumberFormat(20, "h:mm")
        val TIME_HMS = NumberFormat(21, "h:mm:ss")
        val DATETIME = NumberFormat(22, "m/d/yyyy h:mm")

        fun custom(id: Int, formatCode: String) = NumberFormat(id, formatCode)
    }
}
