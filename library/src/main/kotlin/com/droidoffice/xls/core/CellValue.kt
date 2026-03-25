package com.droidoffice.xls.core

import java.time.LocalDate
import java.time.LocalDateTime

/**
 * Represents the typed value of a cell.
 */
sealed class CellValue {
    data object Empty : CellValue()
    data class Text(val value: String) : CellValue()
    data class Number(val value: Double) : CellValue()
    data class Bool(val value: Boolean) : CellValue()
    data class DateValue(val value: LocalDateTime) : CellValue()
    data class Formula(val expression: String, val cachedValue: CellValue? = null) : CellValue()
    data class Error(val code: ErrorCode) : CellValue()
}

enum class ErrorCode(val symbol: String) {
    NULL("#NULL!"),
    DIV_ZERO("#DIV/0!"),
    VALUE("#VALUE!"),
    REF("#REF!"),
    NAME("#NAME?"),
    NUM("#NUM!"),
    NA("#N/A"),
}
