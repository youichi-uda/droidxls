package com.droidoffice.xls.core

/**
 * Represents a cell reference like "A1", "B2", "AA100".
 * Row and column are 0-based internally.
 */
data class CellReference(val row: Int, val col: Int) {

    init {
        require(row >= 0) { "Row must be >= 0, got $row" }
        require(col >= 0) { "Column must be >= 0, got $col" }
    }

    /**
     * Returns the A1-style reference string (e.g. "A1", "B2", "AA100").
     */
    fun toA1(): String = "${columnToLetters(col)}${row + 1}"

    override fun toString(): String = toA1()

    companion object {
        /**
         * Parse an A1-style reference like "A1", "B2", "AA100".
         */
        fun parse(ref: String): CellReference {
            val trimmed = ref.trim().uppercase()
            val colPart = trimmed.takeWhile { it.isLetter() }
            val rowPart = trimmed.dropWhile { it.isLetter() }

            require(colPart.isNotEmpty() && rowPart.isNotEmpty()) {
                "Invalid cell reference: $ref"
            }

            val col = lettersToColumn(colPart)
            val row = rowPart.toIntOrNull()?.minus(1)
                ?: throw IllegalArgumentException("Invalid row in cell reference: $ref")

            return CellReference(row, col)
        }

        fun columnToLetters(col: Int): String {
            val sb = StringBuilder()
            var c = col
            do {
                sb.insert(0, ('A' + c % 26))
                c = c / 26 - 1
            } while (c >= 0)
            return sb.toString()
        }

        fun lettersToColumn(letters: String): Int {
            var col = 0
            for (ch in letters) {
                col = col * 26 + (ch - 'A' + 1)
            }
            return col - 1
        }
    }
}
