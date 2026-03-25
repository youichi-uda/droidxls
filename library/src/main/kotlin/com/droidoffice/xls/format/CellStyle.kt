package com.droidoffice.xls.format

/**
 * Complete style definition for a cell.
 */
data class CellStyle(
    var font: Font = Font(),
    var fill: CellFill = CellFill(),
    var border: BorderSet = BorderSet(),
    var alignment: Alignment = Alignment(),
    var numberFormat: NumberFormat = NumberFormat.GENERAL,
)

/**
 * DSL builder for CellStyle.
 */
class CellStyleBuilder {
    private val style = CellStyle()

    fun font(block: Font.() -> Unit) {
        style.font.apply(block)
    }

    fun fill(block: CellFill.() -> Unit) {
        style.fill.apply(block)
    }

    fun border(block: BorderSet.() -> Unit) {
        style.border.apply(block)
    }

    fun alignment(block: Alignment.() -> Unit) {
        style.alignment.apply(block)
    }

    fun numberFormat(format: NumberFormat) {
        style.numberFormat = format
    }

    internal fun build(): CellStyle = style
}
