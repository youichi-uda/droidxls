package com.droidoffice.xls.exception

import com.droidoffice.core.exception.DroidOfficeException

/**
 * Base exception for DroidXLS-specific errors.
 */
open class DroidXLSException(
    message: String,
    cause: Throwable? = null,
) : DroidOfficeException(message, cause)
