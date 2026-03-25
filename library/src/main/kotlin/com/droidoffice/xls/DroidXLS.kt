package com.droidoffice.xls

import android.content.Context
import android.content.SharedPreferences
import com.droidoffice.core.license.GumroadLicenseVerifier
import com.droidoffice.core.license.LicenseCache
import com.droidoffice.core.license.LicensePlan
import com.droidoffice.core.license.VerifiedLicense
import java.time.LocalDate

/**
 * Entry point for DroidXLS library initialization.
 *
 * ```kotlin
 * // Commercial use: call once in Application.onCreate()
 * DroidXLS.initialize(context, licenseKey = "XXXXXXXX-XXXX-XXXX-XXXX")
 *
 * // Personal/non-commercial use: no initialization needed
 * val workbook = Workbook()
 * ```
 */
object DroidXLS {

    internal var verifier: GumroadLicenseVerifier? = null
        private set

    /**
     * Initialize DroidXLS with a Gumroad license key for commercial use.
     *
     * First call requires network access to verify with Gumroad API.
     * Subsequent calls use cached result (re-verifies every 30 days).
     *
     * For personal/non-commercial use, this call is optional.
     *
     * @param context Android context (for SharedPreferences cache)
     * @param licenseKey License key from Gumroad purchase
     */
    fun initialize(context: Context, licenseKey: String) {
        val cache = SharedPrefsLicenseCache(
            context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        )
        val v = GumroadLicenseVerifier(GUMROAD_PRODUCT_PERMALINK, cache)
        v.initialize(licenseKey)
        verifier = v
    }

    /**
     * Check if the library has been initialized with a license key.
     */
    val isLicensed: Boolean get() = verifier?.isLicensed == true

    /**
     * Get the current license status description.
     */
    val licenseStatus: String
        get() = when {
            verifier == null -> "Personal (no license key)"
            verifier?.isLicensed == true -> "Licensed (${verifier?.currentLicense?.plan})"
            else -> "Not verified"
        }

    internal const val GUMROAD_PRODUCT_PERMALINK = "VF7ciWroOci-4P2kqnw8Wg=="
    private const val PREFS_NAME = "droidxls_license"
}

/**
 * SharedPreferences-backed license cache for Android.
 */
internal class SharedPrefsLicenseCache(
    private val prefs: SharedPreferences,
) : LicenseCache {

    override fun save(license: VerifiedLicense) {
        prefs.edit()
            .putString(KEY_LICENSE_KEY, license.licenseKey)
            .putString(KEY_EMAIL, license.email)
            .putString(KEY_PLAN, license.plan.name)
            .putString(KEY_EXPIRES_AT, license.expiresAt.toString())
            .putString(KEY_VERIFIED_AT, license.verifiedAt.toString())
            .apply()
    }

    override fun load(): VerifiedLicense? {
        val key = prefs.getString(KEY_LICENSE_KEY, null) ?: return null
        val email = prefs.getString(KEY_EMAIL, null) ?: return null
        val planStr = prefs.getString(KEY_PLAN, null) ?: return null
        val expiresStr = prefs.getString(KEY_EXPIRES_AT, null) ?: return null
        val verifiedStr = prefs.getString(KEY_VERIFIED_AT, null) ?: return null

        return try {
            VerifiedLicense(
                licenseKey = key,
                email = email,
                plan = LicensePlan.valueOf(planStr),
                expiresAt = LocalDate.parse(expiresStr),
                verifiedAt = LocalDate.parse(verifiedStr),
            )
        } catch (_: Exception) {
            null
        }
    }

    companion object {
        private const val KEY_LICENSE_KEY = "license_key"
        private const val KEY_EMAIL = "email"
        private const val KEY_PLAN = "plan"
        private const val KEY_EXPIRES_AT = "expires_at"
        private const val KEY_VERIFIED_AT = "verified_at"
    }
}
