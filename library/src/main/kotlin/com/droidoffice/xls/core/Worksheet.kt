package com.droidoffice.xls.core

import com.droidoffice.xls.chart.Chart
import com.droidoffice.xls.chart.ChartBuilder
import com.droidoffice.xls.chart.ChartType
import com.droidoffice.xls.drawing.ImageFormat
import com.droidoffice.xls.drawing.Picture
import com.droidoffice.xls.drawing.pixelsToEmu

/**
 * Represents a single worksheet (sheet) in a workbook.
 */
class Worksheet internal constructor(
    var name: String,
) {
    /** Sparse storage: only non-empty cells are stored. */
    private val cells = mutableMapOf<CellReference, Cell>()

    /** Column width overrides (0-based column index → width in characters). */
    val columnWidths = mutableMapOf<Int, Double>()

    /** Row height overrides (0-based row index → height in points). */
    val rowHeights = mutableMapOf<Int, Double>()

    /** Merged cell ranges. */
    val mergedRegions = mutableListOf<CellRange>()

    /** Hidden columns (0-based). */
    val hiddenColumns = mutableSetOf<Int>()

    /** Hidden rows (0-based). */
    val hiddenRows = mutableSetOf<Int>()

    /** Charts embedded in this sheet. */
    val charts = mutableListOf<Chart>()

    /** Embedded pictures. */
    val pictures = mutableListOf<Picture>()

    /** Auto filter range (e.g. "A1:D10"). */
    var autoFilterRange: CellRange? = null

    /** Data validations. */
    val dataValidations = mutableListOf<DataValidation>()

    /** Conditional formatting rules. */
    val conditionalFormattings = mutableListOf<ConditionalFormatting>()

    /** Sheet protection settings. */
    var protection: SheetProtection? = null

    var isHidden: Boolean = false
    var isActive: Boolean = false

    /** Freeze pane position (null = no freeze). */
    var freezePane: CellReference? = null

    // -- Cell access --

    fun cell(ref: CellReference): Cell =
        cells.getOrPut(ref) { Cell(ref) }

    operator fun get(a1Ref: String): Cell =
        cell(CellReference.parse(a1Ref))

    fun cell(row: Int, col: Int): Cell =
        cell(CellReference(row, col))

    fun cells(): Collection<Cell> = cells.values

    fun getCellOrNull(ref: CellReference): Cell? = cells[ref]

    fun lastRowIndex(): Int = cells.keys.maxOfOrNull { it.row } ?: -1

    fun lastColumnIndex(): Int = cells.keys.maxOfOrNull { it.col } ?: -1

    // -- Row/Column dimensions --

    fun setColumnWidth(col: Int, width: Double) {
        columnWidths[col] = width
    }

    fun setRowHeight(row: Int, height: Double) {
        rowHeights[row] = height
    }

    fun hideColumn(col: Int) { hiddenColumns.add(col) }
    fun showColumn(col: Int) { hiddenColumns.remove(col) }
    fun isColumnHidden(col: Int): Boolean = col in hiddenColumns

    fun hideRow(row: Int) { hiddenRows.add(row) }
    fun showRow(row: Int) { hiddenRows.remove(row) }
    fun isRowHidden(row: Int): Boolean = row in hiddenRows

    // -- Row/Column insert/delete --

    /**
     * Insert [count] empty rows at [rowIndex], shifting existing cells down.
     */
    fun insertRows(rowIndex: Int, count: Int = 1) {
        shiftRows(rowIndex, count)
    }

    /**
     * Delete [count] rows starting at [rowIndex], shifting cells up.
     */
    fun deleteRows(rowIndex: Int, count: Int = 1) {
        // Remove cells in the deleted range
        val toRemove = cells.keys.filter { it.row in rowIndex until rowIndex + count }
        toRemove.forEach { cells.remove(it) }
        // Shift remaining cells up
        shiftRows(rowIndex + count, -count)
    }

    /**
     * Insert [count] empty columns at [colIndex], shifting existing cells right.
     */
    fun insertColumns(colIndex: Int, count: Int = 1) {
        shiftColumns(colIndex, count)
    }

    /**
     * Delete [count] columns starting at [colIndex], shifting cells left.
     */
    fun deleteColumns(colIndex: Int, count: Int = 1) {
        val toRemove = cells.keys.filter { it.col in colIndex until colIndex + count }
        toRemove.forEach { cells.remove(it) }
        shiftColumns(colIndex + count, -count)
    }

    private fun shiftRows(fromRow: Int, delta: Int) {
        val affected = cells.entries
            .filter { it.key.row >= fromRow }
            .sortedBy { if (delta > 0) -it.key.row else it.key.row }

        for ((ref, cell) in affected) {
            cells.remove(ref)
            val newRef = CellReference(ref.row + delta, ref.col)
            val newCell = Cell(newRef)
            newCell.cellValue = cell.cellValue
            newCell.styleIndex = cell.styleIndex
            cells[newRef] = newCell
        }

        // Shift row heights
        val heightEntries = rowHeights.entries.filter { it.key >= fromRow }.toList()
        for ((row, height) in heightEntries) {
            rowHeights.remove(row)
            val newRow = row + delta
            if (newRow >= 0) rowHeights[newRow] = height
        }

        // Shift hidden rows
        val hiddenToShift = hiddenRows.filter { it >= fromRow }.toList()
        hiddenRows.removeAll(hiddenToShift.toSet())
        hiddenRows.addAll(hiddenToShift.map { it + delta }.filter { it >= 0 })
    }

    private fun shiftColumns(fromCol: Int, delta: Int) {
        val affected = cells.entries
            .filter { it.key.col >= fromCol }
            .sortedBy { if (delta > 0) -it.key.col else it.key.col }

        for ((ref, cell) in affected) {
            cells.remove(ref)
            val newRef = CellReference(ref.row, ref.col + delta)
            val newCell = Cell(newRef)
            newCell.cellValue = cell.cellValue
            newCell.styleIndex = cell.styleIndex
            cells[newRef] = newCell
        }

        val widthEntries = columnWidths.entries.filter { it.key >= fromCol }.toList()
        for ((col, width) in widthEntries) {
            columnWidths.remove(col)
            val newCol = col + delta
            if (newCol >= 0) columnWidths[newCol] = width
        }

        val hiddenToShift = hiddenColumns.filter { it >= fromCol }.toList()
        hiddenColumns.removeAll(hiddenToShift.toSet())
        hiddenColumns.addAll(hiddenToShift.map { it + delta }.filter { it >= 0 })
    }

    // -- Merge/Freeze --

    fun addMergedRegion(range: CellRange) {
        mergedRegions.add(range)
    }

    fun removeMergedRegion(index: Int) {
        mergedRegions.removeAt(index)
    }

    fun freeze(row: Int, col: Int) {
        freezePane = CellReference(row, col)
    }

    fun unfreeze() {
        freezePane = null
    }

    // -- Pictures --

    fun addPicture(
        imageData: ByteArray,
        format: ImageFormat,
        fromCol: Int, fromRow: Int,
        toCol: Int, toRow: Int,
        widthPx: Int = 0, heightPx: Int = 0,
    ): Picture {
        val pic = Picture(
            imageData = imageData, format = format,
            fromCol = fromCol, fromRow = fromRow,
            toCol = toCol, toRow = toRow,
            widthEmu = pixelsToEmu(widthPx), heightEmu = pixelsToEmu(heightPx),
        )
        pictures.add(pic)
        return pic
    }

    // -- Charts --

    fun addChart(type: ChartType, block: ChartBuilder.() -> Unit): Chart {
        val builder = ChartBuilder(type)
        builder.apply(block)
        val chart = builder.build()
        charts.add(chart)
        return chart
    }

    // -- Protection --

    fun protect(password: String? = null) {
        protection = if (password != null) SheetProtection.withPassword(password) else SheetProtection.noPassword()
    }

    fun unprotect() {
        protection = null
    }
}
