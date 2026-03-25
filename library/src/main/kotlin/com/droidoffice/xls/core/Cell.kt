package com.droidoffice.xls.core

import com.droidoffice.xls.format.CellStyle
import com.droidoffice.xls.format.CellStyleBuilder
import java.time.LocalDate
import java.time.LocalDateTime

/**
 * Represents a single cell in a worksheet.
 */
class Cell internal constructor(
    val reference: CellReference,
) {
    var cellValue: CellValue = CellValue.Empty
        internal set

    /** Resolved style (populated during write or set via DSL). */
    var cellStyle: CellStyle? = null

    /** Style index referencing the workbook's style table. */
    var styleIndex: Int = 0

    // -- Convenience getters/setters --

    var value: Any?
        get() = when (val v = cellValue) {
            is CellValue.Empty -> null
            is CellValue.Text -> v.value
            is CellValue.Number -> v.value
            is CellValue.Bool -> v.value
            is CellValue.DateValue -> v.value
            is CellValue.Formula -> v.cachedValue?.let { value }
            is CellValue.Error -> v.code.symbol
        }
        set(v) {
            cellValue = when (v) {
                null -> CellValue.Empty
                is String -> CellValue.Text(v)
                is Int -> CellValue.Number(v.toDouble())
                is Long -> CellValue.Number(v.toDouble())
                is Float -> CellValue.Number(v.toDouble())
                is Double -> CellValue.Number(v)
                is Boolean -> CellValue.Bool(v)
                is LocalDate -> CellValue.DateValue(v.atStartOfDay())
                is LocalDateTime -> CellValue.DateValue(v)
                else -> CellValue.Text(v.toString())
            }
        }

    var formula: String?
        get() = (cellValue as? CellValue.Formula)?.expression
        set(expr) {
            cellValue = if (expr != null) {
                CellValue.Formula(expr.removePrefix("="))
            } else {
                CellValue.Empty
            }
        }

    val stringValue: String
        get() = when (val v = cellValue) {
            is CellValue.Empty -> ""
            is CellValue.Text -> v.value
            is CellValue.Number -> {
                if (v.value == v.value.toLong().toDouble()) {
                    v.value.toLong().toString()
                } else {
                    v.value.toString()
                }
            }
            is CellValue.Bool -> v.value.toString()
            is CellValue.DateValue -> v.value.toString()
            is CellValue.Formula -> v.cachedValue?.let { Cell(reference).apply { cellValue = it }.stringValue } ?: ""
            is CellValue.Error -> v.code.symbol
        }

    val numericValue: Double
        get() = when (val v = cellValue) {
            is CellValue.Number -> v.value
            is CellValue.Bool -> if (v.value) 1.0 else 0.0
            else -> 0.0
        }

    val isEmpty: Boolean get() = cellValue is CellValue.Empty

    /**
     * Apply a style using Kotlin DSL.
     * ```
     * cell.style {
     *     font { bold = true; size = 14.0 }
     *     fill { backgroundColor = OfficeColor.LIGHT_BLUE }
     *     border { all = BorderStyle.THIN }
     * }
     * ```
     */
    fun style(block: CellStyleBuilder.() -> Unit) {
        val builder = CellStyleBuilder()
        // Start from existing style if present
        cellStyle?.let {
            builder.font { this.name = it.font.name; this.size = it.font.size; this.bold = it.font.bold; this.italic = it.font.italic; this.underline = it.font.underline; this.strikethrough = it.font.strikethrough; this.color = it.font.color }
            builder.fill { this.patternType = it.fill.patternType; this.foregroundColor = it.fill.foregroundColor; this.backgroundColor = it.fill.backgroundColor }
            builder.border { this.left = it.border.left.copy(); this.right = it.border.right.copy(); this.top = it.border.top.copy(); this.bottom = it.border.bottom.copy() }
            builder.alignment { this.horizontal = it.alignment.horizontal; this.vertical = it.alignment.vertical; this.wrapText = it.alignment.wrapText; this.indent = it.alignment.indent; this.rotation = it.alignment.rotation }
            builder.numberFormat(it.numberFormat)
        }
        builder.apply(block)
        cellStyle = builder.build()
    }
}
