package com.droidoffice.xls.core

/**
 * Represents a rectangular range of cells (e.g. "A1:C3").
 */
data class CellRange(
    val topLeft: CellReference,
    val bottomRight: CellReference,
) {
    val firstRow: Int get() = topLeft.row
    val lastRow: Int get() = bottomRight.row
    val firstCol: Int get() = topLeft.col
    val lastCol: Int get() = bottomRight.col

    fun toA1(): String = "${topLeft.toA1()}:${bottomRight.toA1()}"

    override fun toString(): String = toA1()

    companion object {
        fun parse(range: String): CellRange {
            val parts = range.split(":")
            require(parts.size == 2) { "Invalid range: $range" }
            return CellRange(
                CellReference.parse(parts[0]),
                CellReference.parse(parts[1]),
            )
        }
    }
}
