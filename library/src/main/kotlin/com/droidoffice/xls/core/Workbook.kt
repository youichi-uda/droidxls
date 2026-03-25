package com.droidoffice.xls.core

import com.droidoffice.xls.format.StyleSheet
import com.droidoffice.xls.io.XlsxReader
import com.droidoffice.xls.io.XlsxWriter
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.InputStream
import java.io.OutputStream

/**
 * The main entry point for working with Excel workbooks.
 */
class Workbook {

    private val _sheets = mutableListOf<Worksheet>()

    /** All worksheets in this workbook. */
    val sheets: List<Worksheet> get() = _sheets

    /** Shared string table (populated during read). */
    internal val sharedStrings = mutableListOf<String>()

    /** Style sheet (populated during read, used for date detection etc.). */
    internal val styleSheet = StyleSheet()

    /** Named ranges. */
    val namedRanges = mutableListOf<NamedRange>()

    // -- Sheet operations --

    fun addSheet(name: String): Worksheet {
        val sheet = Worksheet(name)
        if (_sheets.isEmpty()) sheet.isActive = true
        _sheets.add(sheet)
        return sheet
    }

    fun insertSheet(index: Int, name: String): Worksheet {
        val sheet = Worksheet(name)
        _sheets.add(index, sheet)
        return sheet
    }

    fun removeSheet(index: Int) {
        _sheets.removeAt(index)
    }

    fun removeSheet(name: String) {
        _sheets.removeAll { it.name == name }
    }

    fun getSheet(name: String): Worksheet? =
        _sheets.find { it.name == name }

    fun getSheet(index: Int): Worksheet = _sheets[index]

    fun moveSheet(fromIndex: Int, toIndex: Int) {
        val sheet = _sheets.removeAt(fromIndex)
        _sheets.add(toIndex, sheet)
    }

    fun copySheet(sourceIndex: Int, newName: String): Worksheet {
        val source = _sheets[sourceIndex]
        val copy = Worksheet(newName)
        copy.isHidden = source.isHidden
        copy.freezePane = source.freezePane
        copy.columnWidths.putAll(source.columnWidths)
        copy.rowHeights.putAll(source.rowHeights)
        copy.hiddenColumns.addAll(source.hiddenColumns)
        copy.hiddenRows.addAll(source.hiddenRows)
        for (region in source.mergedRegions) {
            copy.mergedRegions.add(region)
        }
        for (cell in source.cells()) {
            val newCell = copy.cell(cell.reference)
            newCell.cellValue = cell.cellValue
            newCell.styleIndex = cell.styleIndex
        }
        _sheets.add(copy)
        return copy
    }

    val sheetCount: Int get() = _sheets.size

    var activeSheetIndex: Int
        get() = _sheets.indexOfFirst { it.isActive }.takeIf { it >= 0 } ?: 0
        set(index) {
            _sheets.forEach { it.isActive = false }
            if (index in _sheets.indices) _sheets[index].isActive = true
        }

    val activeSheet: Worksheet?
        get() = _sheets.find { it.isActive } ?: _sheets.firstOrNull()

    // -- I/O --

    fun save(output: OutputStream) {
        XlsxWriter.write(this, output)
    }

    fun save(output: OutputStream, password: String) {
        val buffer = java.io.ByteArrayOutputStream()
        XlsxWriter.write(this, buffer)
        val encrypted = com.droidoffice.core.ooxml.EncryptedPackage.encrypt(buffer.toByteArray(), password)
        output.write(encrypted)
    }

    suspend fun saveAsync(output: OutputStream) {
        withContext(Dispatchers.IO) { save(output) }
    }

    suspend fun saveAsync(output: OutputStream, password: String) {
        withContext(Dispatchers.IO) { save(output, password) }
    }

    companion object {
        fun open(input: InputStream): Workbook {
            val bytes = input.readBytes()
            return if (com.droidoffice.core.ooxml.EncryptedPackage.isEncrypted(bytes)) {
                throw com.droidoffice.core.exception.PasswordException(
                    "This file is password-protected. Use open(input, password) instead."
                )
            } else {
                XlsxReader.read(java.io.ByteArrayInputStream(bytes))
            }
        }

        fun open(input: InputStream, password: String): Workbook {
            val bytes = input.readBytes()
            val decrypted = if (com.droidoffice.core.ooxml.EncryptedPackage.isEncrypted(bytes)) {
                com.droidoffice.core.ooxml.EncryptedPackage.decrypt(bytes, password)
            } else {
                bytes // Not encrypted, just read normally
            }
            return XlsxReader.read(java.io.ByteArrayInputStream(decrypted))
        }

        suspend fun openAsync(input: InputStream): Workbook {
            return withContext(Dispatchers.IO) { open(input) }
        }

        suspend fun openAsync(input: InputStream, password: String): Workbook {
            return withContext(Dispatchers.IO) { open(input, password) }
        }
    }
}
