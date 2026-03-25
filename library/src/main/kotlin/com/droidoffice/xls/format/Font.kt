package com.droidoffice.xls.format

import com.droidoffice.core.drawingml.OfficeColor
import com.droidoffice.core.drawingml.UnderlineStyle

/**
 * Font properties for a cell style.
 */
data class Font(
    var name: String = "Calibri",
    var size: Double = 11.0,
    var bold: Boolean = false,
    var italic: Boolean = false,
    var underline: UnderlineStyle = UnderlineStyle.NONE,
    var strikethrough: Boolean = false,
    var color: OfficeColor? = null,
) {
    internal fun isDefault(): Boolean =
        name == "Calibri" && size == 11.0 && !bold && !italic &&
            underline == UnderlineStyle.NONE && !strikethrough && color == null
}
