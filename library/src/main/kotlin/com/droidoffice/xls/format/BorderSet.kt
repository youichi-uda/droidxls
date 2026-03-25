package com.droidoffice.xls.format

import com.droidoffice.core.drawingml.BorderStyle
import com.droidoffice.core.drawingml.OfficeColor

data class CellBorder(
    var style: BorderStyle = BorderStyle.NONE,
    var color: OfficeColor? = null,
) {
    internal fun isDefault(): Boolean = style == BorderStyle.NONE && color == null
}

/**
 * All four borders of a cell.
 */
data class BorderSet(
    var left: CellBorder = CellBorder(),
    var right: CellBorder = CellBorder(),
    var top: CellBorder = CellBorder(),
    var bottom: CellBorder = CellBorder(),
) {
    /** Set all borders at once. */
    var all: BorderStyle
        get() = left.style // arbitrary
        set(value) {
            left.style = value; right.style = value
            top.style = value; bottom.style = value
        }

    var allColor: OfficeColor?
        get() = left.color
        set(value) {
            left.color = value; right.color = value
            top.color = value; bottom.color = value
        }

    internal fun isDefault(): Boolean =
        left.isDefault() && right.isDefault() && top.isDefault() && bottom.isDefault()
}
