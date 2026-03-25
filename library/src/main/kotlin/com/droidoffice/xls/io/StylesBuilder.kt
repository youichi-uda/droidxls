package com.droidoffice.xls.io

import com.droidoffice.core.drawingml.BorderStyle
import com.droidoffice.core.drawingml.OfficeColor
import com.droidoffice.core.drawingml.PatternType
import com.droidoffice.core.drawingml.UnderlineStyle
import com.droidoffice.xls.format.*

/**
 * Collects unique styles from all cells and builds the styles.xml content
 * along with a mapping from CellStyle to style index.
 */
class StylesBuilder {

    private val fonts = mutableListOf<Font>()
    private val fills = mutableListOf<CellFill>()
    private val borders = mutableListOf<BorderSet>()
    private val numFmts = mutableMapOf<Int, String>()
    private val xfs = mutableListOf<XfEntry>()
    private val styleToIndex = mutableMapOf<CellStyle, Int>()

    init {
        // Default font (index 0)
        fonts.add(Font())
        // Required fills: none (0) and gray125 (1)
        fills.add(CellFill(PatternType.NONE))
        fills.add(CellFill(PatternType.GRAY_125))
        // Default border (index 0)
        borders.add(BorderSet())
        // Default xf (index 0)
        xfs.add(XfEntry(0, 0, 0, 0, Alignment()))
        styleToIndex[CellStyle()] = 0
    }

    /**
     * Register a CellStyle and return its style index.
     */
    fun registerStyle(style: CellStyle?): Int {
        if (style == null) return 0
        styleToIndex[style]?.let { return it }

        val fontIdx = registerFont(style.font)
        val fillIdx = registerFill(style.fill)
        val borderIdx = registerBorder(style.border)
        val numFmtId = registerNumFmt(style.numberFormat)

        val xfIdx = xfs.size
        xfs.add(XfEntry(numFmtId, fontIdx, fillIdx, borderIdx, style.alignment))
        styleToIndex[style] = xfIdx
        return xfIdx
    }

    private fun registerFont(font: Font): Int {
        val existing = fonts.indexOf(font)
        if (existing >= 0) return existing
        fonts.add(font)
        return fonts.size - 1
    }

    private fun registerFill(fill: CellFill): Int {
        val existing = fills.indexOf(fill)
        if (existing >= 0) return existing
        fills.add(fill)
        return fills.size - 1
    }

    private fun registerBorder(border: BorderSet): Int {
        val existing = borders.indexOf(border)
        if (existing >= 0) return existing
        borders.add(border)
        return borders.size - 1
    }

    private fun registerNumFmt(numFmt: NumberFormat): Int {
        if (numFmt.id < 164) return numFmt.id // Built-in
        numFmts[numFmt.id] = numFmt.formatCode
        return numFmt.id
    }

    fun buildStylesXml(): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""")

        // numFmts
        if (numFmts.isNotEmpty()) {
            appendLine("  <numFmts count=\"${numFmts.size}\">")
            for ((id, code) in numFmts) {
                appendLine("    <numFmt numFmtId=\"$id\" formatCode=\"${escapeXml(code)}\"/>")
            }
            appendLine("  </numFmts>")
        }

        // fonts
        appendLine("  <fonts count=\"${fonts.size}\">")
        for (font in fonts) {
            appendLine("    <font>")
            if (font.bold) appendLine("      <b/>")
            if (font.italic) appendLine("      <i/>")
            if (font.strikethrough) appendLine("      <strike/>")
            if (font.underline != UnderlineStyle.NONE) {
                val uVal = when (font.underline) {
                    UnderlineStyle.SINGLE -> "single"
                    UnderlineStyle.DOUBLE -> "double"
                    UnderlineStyle.SINGLE_ACCOUNTING -> "singleAccounting"
                    UnderlineStyle.DOUBLE_ACCOUNTING -> "doubleAccounting"
                    else -> "single"
                }
                appendLine("      <u val=\"$uVal\"/>")
            }
            appendLine("      <sz val=\"${font.size}\"/>")
            font.color?.let { writeColor(it, "      ") }
            appendLine("      <name val=\"${escapeXml(font.name)}\"/>")
            appendLine("    </font>")
        }
        appendLine("  </fonts>")

        // fills
        appendLine("  <fills count=\"${fills.size}\">")
        for (fill in fills) {
            appendLine("    <fill>")
            val patternStr = patternTypeToString(fill.patternType)
            append("      <patternFill patternType=\"$patternStr\"")
            if (fill.foregroundColor != null || fill.backgroundColor != null) {
                appendLine(">")
                fill.foregroundColor?.let {
                    append("        <fgColor ")
                    appendColorAttr(it)
                    appendLine("/>")
                }
                fill.backgroundColor?.let {
                    append("        <bgColor ")
                    appendColorAttr(it)
                    appendLine("/>")
                }
                appendLine("      </patternFill>")
            } else {
                appendLine("/>")
            }
            appendLine("    </fill>")
        }
        appendLine("  </fills>")

        // borders
        appendLine("  <borders count=\"${borders.size}\">")
        for (border in borders) {
            appendLine("    <border>")
            writeBorderSide("left", border.left)
            writeBorderSide("right", border.right)
            writeBorderSide("top", border.top)
            writeBorderSide("bottom", border.bottom)
            appendLine("      <diagonal/>")
            appendLine("    </border>")
        }
        appendLine("  </borders>")

        // cellStyleXfs
        appendLine("  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>")

        // cellXfs
        appendLine("  <cellXfs count=\"${xfs.size}\">")
        for (xf in xfs) {
            append("    <xf numFmtId=\"${xf.numFmtId}\" fontId=\"${xf.fontId}\" fillId=\"${xf.fillId}\" borderId=\"${xf.borderId}\" xfId=\"0\"")
            if (xf.numFmtId != 0) append(" applyNumberFormat=\"1\"")
            if (xf.fontId != 0) append(" applyFont=\"1\"")
            if (xf.fillId != 0) append(" applyFill=\"1\"")
            if (xf.borderId != 0) append(" applyBorder=\"1\"")
            if (!xf.alignment.isDefault()) {
                appendLine(" applyAlignment=\"1\">")
                append("      <alignment")
                if (xf.alignment.horizontal != HorizontalAlignment.GENERAL)
                    append(" horizontal=\"${xf.alignment.horizontal.name.lowercase()}\"")
                if (xf.alignment.vertical != VerticalAlignment.BOTTOM)
                    append(" vertical=\"${xf.alignment.vertical.name.lowercase()}\"")
                if (xf.alignment.wrapText) append(" wrapText=\"1\"")
                if (xf.alignment.indent > 0) append(" indent=\"${xf.alignment.indent}\"")
                if (xf.alignment.rotation != 0) append(" textRotation=\"${xf.alignment.rotation}\"")
                appendLine("/>")
                appendLine("    </xf>")
            } else {
                appendLine("/>")
            }
        }
        appendLine("  </cellXfs>")

        appendLine("</styleSheet>")
    }

    private fun StringBuilder.writeColor(color: OfficeColor, indent: String) {
        append("${indent}<color ")
        appendColorAttr(color)
        appendLine("/>")
    }

    private fun StringBuilder.appendColorAttr(color: OfficeColor) {
        when (color) {
            is OfficeColor.Rgb -> append("rgb=\"FF${color.toHex()}\"")
            is OfficeColor.Theme -> {
                append("theme=\"${color.themeIndex}\"")
                if (color.tint != 0.0) append(" tint=\"${color.tint}\"")
            }
            is OfficeColor.Indexed -> append("indexed=\"${color.index}\"")
            is OfficeColor.Auto -> append("auto=\"1\"")
        }
    }

    private fun StringBuilder.writeBorderSide(side: String, border: CellBorder) {
        if (border.isDefault()) {
            appendLine("      <$side/>")
        } else {
            val styleStr = borderStyleToString(border.style)
            appendLine("      <$side style=\"$styleStr\">")
            border.color?.let {
                append("        <color ")
                appendColorAttr(it)
                appendLine("/>")
            }
            appendLine("      </$side>")
        }
    }

    private data class XfEntry(
        val numFmtId: Int,
        val fontId: Int,
        val fillId: Int,
        val borderId: Int,
        val alignment: Alignment,
    )

    companion object {
        private fun patternTypeToString(type: PatternType): String = when (type) {
            PatternType.NONE -> "none"
            PatternType.SOLID -> "solid"
            PatternType.GRAY_125 -> "gray125"
            PatternType.GRAY_0625 -> "gray0625"
            PatternType.DARK_GRAY -> "darkGray"
            PatternType.MEDIUM_GRAY -> "mediumGray"
            PatternType.LIGHT_GRAY -> "lightGray"
            else -> type.name.lowercase().replace("_", "")
        }

        private fun borderStyleToString(style: BorderStyle): String = when (style) {
            BorderStyle.NONE -> "none"
            BorderStyle.THIN -> "thin"
            BorderStyle.MEDIUM -> "medium"
            BorderStyle.THICK -> "thick"
            BorderStyle.DASHED -> "dashed"
            BorderStyle.DOTTED -> "dotted"
            BorderStyle.DOUBLE -> "double"
            BorderStyle.HAIR -> "hair"
            BorderStyle.MEDIUM_DASHED -> "mediumDashed"
            BorderStyle.DASH_DOT -> "dashDot"
            BorderStyle.MEDIUM_DASH_DOT -> "mediumDashDot"
            BorderStyle.DASH_DOT_DOT -> "dashDotDot"
            BorderStyle.MEDIUM_DASH_DOT_DOT -> "mediumDashDotDot"
            BorderStyle.SLANT_DASH_DOT -> "slantDashDot"
        }

        private fun escapeXml(text: String): String = text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\"", "&quot;")
    }
}
