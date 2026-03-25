package com.droidoffice.xls.core

import com.droidoffice.core.drawingml.OfficeColor

/**
 * Conditional formatting applied to a range.
 */
data class ConditionalFormatting(
    val ranges: List<CellRange>,
    val rules: List<ConditionalRule>,
)

data class ConditionalRule(
    val type: ConditionalType,
    val operator: ConditionalOperator = ConditionalOperator.NONE,
    val formula: String? = null,
    val formula2: String? = null,
    val priority: Int = 1,
    val fontBold: Boolean? = null,
    val fontColor: OfficeColor? = null,
    val fillColor: OfficeColor? = null,
)

enum class ConditionalType {
    CELL_IS, EXPRESSION, COLOR_SCALE, DATA_BAR, ICON_SET,
    TOP_10, ABOVE_AVERAGE, DUPLICATE_VALUES, UNIQUE_VALUES,
    CONTAINS_TEXT, NOT_CONTAINS_TEXT, BEGINS_WITH, ENDS_WITH,
}

enum class ConditionalOperator {
    NONE, BETWEEN, NOT_BETWEEN, EQUAL, NOT_EQUAL,
    GREATER_THAN, LESS_THAN, GREATER_THAN_OR_EQUAL, LESS_THAN_OR_EQUAL,
}
