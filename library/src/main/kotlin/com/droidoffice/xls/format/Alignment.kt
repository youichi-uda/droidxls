package com.droidoffice.xls.format

enum class HorizontalAlignment {
    GENERAL, LEFT, CENTER, RIGHT, FILL, JUSTIFY, CENTER_CONTINUOUS, DISTRIBUTED
}

enum class VerticalAlignment {
    TOP, CENTER, BOTTOM, JUSTIFY, DISTRIBUTED
}

data class Alignment(
    var horizontal: HorizontalAlignment = HorizontalAlignment.GENERAL,
    var vertical: VerticalAlignment = VerticalAlignment.BOTTOM,
    var wrapText: Boolean = false,
    var indent: Int = 0,
    var rotation: Int = 0,
) {
    internal fun isDefault(): Boolean =
        horizontal == HorizontalAlignment.GENERAL && vertical == VerticalAlignment.BOTTOM &&
            !wrapText && indent == 0 && rotation == 0
}
