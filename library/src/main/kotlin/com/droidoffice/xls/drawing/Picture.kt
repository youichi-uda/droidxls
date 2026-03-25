package com.droidoffice.xls.drawing

/**
 * Represents an image embedded in a worksheet.
 */
data class Picture(
    val imageData: ByteArray,
    val format: ImageFormat,
    val fromCol: Int,
    val fromRow: Int,
    val toCol: Int,
    val toRow: Int,
    val widthEmu: Long = 0,
    val heightEmu: Long = 0,
) {
    /** Internal media path (e.g. "xl/media/image1.png") set during write. */
    internal var mediaPath: String = ""
    internal var rId: String = ""

    override fun equals(other: Any?): Boolean {
        if (this === other) return true
        if (other !is Picture) return false
        return format == other.format && fromCol == other.fromCol && fromRow == other.fromRow &&
            toCol == other.toCol && toRow == other.toRow && imageData.contentEquals(other.imageData)
    }

    override fun hashCode(): Int = imageData.contentHashCode()
}

enum class ImageFormat(val extension: String, val contentType: String) {
    PNG("png", "image/png"),
    JPEG("jpeg", "image/jpeg"),
    WEBP("webp", "image/webp"),
}

/** EMU (English Metric Units) per pixel at 96 DPI */
const val EMU_PER_PIXEL = 9525L

fun pixelsToEmu(pixels: Int): Long = pixels * EMU_PER_PIXEL
