package com.droidoffice.xls.core

/**
 * A named range definition (e.g. "SalesData" referring to "Sheet1!$A$1:$D$100").
 */
data class NamedRange(
    val name: String,
    val reference: String,
    val sheetIndex: Int = -1, // -1 = workbook-scoped
)
