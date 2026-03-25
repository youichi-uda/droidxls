package com.droidoffice.xls

import android.content.Context
import com.droidoffice.core.license.LicenseKeyGenerator
import com.droidoffice.core.license.LicenseValidator
import java.security.interfaces.RSAPublicKey

/**
 * Entry point for DroidXLS library initialization.
 *
 * ```kotlin
 * // Commercial use: call once in Application.onCreate()
 * DroidXLS.initialize(context, licenseKey = "your-jwt-license-key")
 *
 * // Personal/non-commercial use: no initialization needed
 * val workbook = Workbook()
 * ```
 */
object DroidXLS {

    internal var licenseValidator: LicenseValidator? = null
        private set

    /**
     * Initialize DroidXLS with a license key for commercial use.
     * Call this once in Application.onCreate() or before any DroidXLS operations.
     *
     * For personal/non-commercial use, this call is optional.
     *
     * @param context Android context (used for future features, not stored)
     * @param licenseKey JWT license key obtained from Gumroad purchase
     */
    fun initialize(context: Context, licenseKey: String) {
        val publicKey = loadPublicKey()
        val validator = LicenseValidator(publicKey, PRODUCT_ID)
        validator.initialize(licenseKey)
        licenseValidator = validator
    }

    /**
     * Check if the library has been initialized with a license key.
     */
    val isLicensed: Boolean get() = licenseValidator != null

    /**
     * Get the current license status description.
     */
    val licenseStatus: String
        get() = when {
            licenseValidator == null -> "Personal (no license key)"
            else -> "Licensed"
        }

    private fun loadPublicKey(): RSAPublicKey {
        // The public key is embedded at build time. In development, use the placeholder.
        // For production release, replace EMBEDDED_PUBLIC_KEY with the actual Base64-encoded key.
        if (EMBEDDED_PUBLIC_KEY.isEmpty()) {
            throw IllegalStateException(
                "License validation not configured. " +
                    "DroidXLS is free for personal/non-commercial use without initialization."
            )
        }
        return LicenseKeyGenerator.decodePublicKey(EMBEDDED_PUBLIC_KEY)
    }

    private const val PRODUCT_ID = "droidxls"

    // Replace with actual public key before production release.
    // Generate with: LicenseKeyGenerator.generateKeyPair() + encodePublicKey()
    internal const val EMBEDDED_PUBLIC_KEY = ""
}
