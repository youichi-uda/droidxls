package com.droidoffice.xls.e2e

import com.droidoffice.core.license.GumroadLicenseVerifier
import com.droidoffice.core.license.LicensePlan
import org.junit.jupiter.api.Assumptions
import org.junit.jupiter.api.Test
import kotlin.test.assertEquals
import kotlin.test.assertNotNull
import kotlin.test.assertTrue

/**
 * Integration test: verifies a real license key against Gumroad API.
 * Requires network access and GUMROAD_LICENSE_KEY environment variable.
 *
 * Skipped in CI (no license key set).
 * To run locally: export GUMROAD_LICENSE_KEY=XXXXXXXX-XXXX-XXXX-XXXX
 */
class GumroadLicenseVerifyTest {

    @Test
    fun `verify real Gumroad license key`() {
        val licenseKey = System.getenv("GUMROAD_LICENSE_KEY")
        Assumptions.assumeTrue(
            !licenseKey.isNullOrBlank(),
            "Skipped: GUMROAD_LICENSE_KEY not set"
        )

        val verifier = GumroadLicenseVerifier(
            productId = "VF7ciWroOci-4P2kqnw8Wg==",
            cache = null,
        )

        verifier.initialize(licenseKey)

        assertTrue(verifier.isLicensed)
        val license = verifier.currentLicense
        assertNotNull(license)
        assertEquals(LicensePlan.COMMERCIAL, license.plan)
        assertTrue(license.email.isNotEmpty())

        // checkLicense should not throw
        verifier.checkLicense()
    }
}
