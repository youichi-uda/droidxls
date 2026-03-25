package com.droidoffice.xls.core

import java.security.MessageDigest

/**
 * Sheet protection settings.
 */
data class SheetProtection(
    val passwordHash: String? = null,
    val sheet: Boolean = true,
    val objects: Boolean = true,
    val scenarios: Boolean = true,
    val formatCells: Boolean = false,
    val formatColumns: Boolean = false,
    val formatRows: Boolean = false,
    val insertColumns: Boolean = false,
    val insertRows: Boolean = false,
    val insertHyperlinks: Boolean = false,
    val deleteColumns: Boolean = false,
    val deleteRows: Boolean = false,
    val selectLockedCells: Boolean = false,
    val sort: Boolean = false,
    val autoFilter: Boolean = false,
    val pivotTables: Boolean = false,
    val selectUnlockedCells: Boolean = false,
) {
    companion object {
        /**
         * Create protection with a password.
         * Uses the Excel legacy hash algorithm for compatibility.
         */
        fun withPassword(password: String): SheetProtection {
            return SheetProtection(passwordHash = hashPassword(password))
        }

        /**
         * Create protection without a password.
         */
        fun noPassword(): SheetProtection = SheetProtection()

        internal fun hashPassword(password: String): String {
            // Excel legacy password hash (simple XOR-based)
            var hash = 0
            for ((i, ch) in password.reversed().withIndex()) {
                var charValue = ch.code
                charValue = ((charValue shr 1) or ((charValue and 1) shl 14)) and 0x7FFF
                hash = hash xor charValue
                hash = ((hash shr 1) or ((hash and 1) shl 14)) and 0x7FFF
            }
            hash = hash xor password.length
            hash = hash xor 0xCE4B
            return hash.toString(16).uppercase()
        }
    }
}
