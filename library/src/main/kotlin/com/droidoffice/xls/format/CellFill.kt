package com.droidoffice.xls.format

import com.droidoffice.core.drawingml.OfficeColor
import com.droidoffice.core.drawingml.PatternType

data class CellFill(
    var patternType: PatternType = PatternType.NONE,
    var foregroundColor: OfficeColor? = null,
    var backgroundColor: OfficeColor? = null,
) {
    internal fun isDefault(): Boolean = patternType == PatternType.NONE
}
