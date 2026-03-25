package com.droidoffice.xls.core

/**
 * Data validation rule applied to a range of cells.
 */
data class DataValidation(
    val range: CellRange,
    val type: ValidationType = ValidationType.NONE,
    val operator: ValidationOperator = ValidationOperator.BETWEEN,
    val formula1: String? = null,
    val formula2: String? = null,
    val showDropDown: Boolean = true,
    val showErrorMessage: Boolean = true,
    val errorTitle: String? = null,
    val errorMessage: String? = null,
    val showInputMessage: Boolean = false,
    val promptTitle: String? = null,
    val promptMessage: String? = null,
)

enum class ValidationType {
    NONE, WHOLE, DECIMAL, LIST, DATE, TIME, TEXT_LENGTH, CUSTOM
}

enum class ValidationOperator {
    BETWEEN, NOT_BETWEEN, EQUAL, NOT_EQUAL,
    GREATER_THAN, LESS_THAN, GREATER_THAN_OR_EQUAL, LESS_THAN_OR_EQUAL,
}
