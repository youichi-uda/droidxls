package com.droidoffice.xls.io

import com.droidoffice.core.ooxml.OoxmlPackage
import com.droidoffice.xls.drawing.Picture

/**
 * Writes Drawing XML parts for embedded images.
 */
object DrawingWriter {

    fun writeDrawings(
        pkg: OoxmlPackage,
        sheetIndex: Int,
        pictures: List<Picture>,
    ) {
        if (pictures.isEmpty()) return

        val drawingIdx = sheetIndex + 1

        // Write media files
        for ((i, pic) in pictures.withIndex()) {
            val mediaPath = "xl/media/image${drawingIdx}_${i + 1}.${pic.format.extension}"
            pkg.setPart(mediaPath, pic.imageData)
            pic.mediaPath = mediaPath
            pic.rId = "rId${i + 1}"
        }

        // Write drawing rels
        pkg.setPart(
            "xl/drawings/_rels/drawing$drawingIdx.xml.rels",
            buildDrawingRels(pictures).toByteArray()
        )

        // Write drawing XML
        pkg.setPart(
            "xl/drawings/drawing$drawingIdx.xml",
            buildDrawingXml(pictures).toByteArray()
        )

        // Write sheet rels for drawing reference
        val sheetRelsPath = "xl/worksheets/_rels/sheet$drawingIdx.xml.rels"
        val existingRels = pkg.getPart(sheetRelsPath)?.decodeToString() ?: ""
        val drawingRel = buildSheetDrawingRel(drawingIdx)
        pkg.setPart(sheetRelsPath, drawingRel.toByteArray())
    }

    private fun buildDrawingRels(pictures: List<Picture>): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""")
        for (pic in pictures) {
            val target = "../${pic.mediaPath.removePrefix("xl/")}"
            appendLine("""  <Relationship Id="${pic.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="$target"/>""")
        }
        appendLine("</Relationships>")
    }

    private fun buildDrawingXml(pictures: List<Picture>): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""")

        for ((i, pic) in pictures.withIndex()) {
            appendLine("  <xdr:twoCellAnchor editAs=\"oneCell\">")
            appendLine("    <xdr:from><xdr:col>${pic.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${pic.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>")
            appendLine("    <xdr:to><xdr:col>${pic.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${pic.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>")
            appendLine("    <xdr:pic>")
            appendLine("      <xdr:nvPicPr>")
            appendLine("        <xdr:cNvPr id=\"${i + 2}\" name=\"Picture ${i + 1}\"/>")
            appendLine("        <xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>")
            appendLine("      </xdr:nvPicPr>")
            appendLine("      <xdr:blipFill>")
            appendLine("        <a:blip r:embed=\"${pic.rId}\"/>")
            appendLine("        <a:stretch><a:fillRect/></a:stretch>")
            appendLine("      </xdr:blipFill>")
            appendLine("      <xdr:spPr>")
            if (pic.widthEmu > 0 && pic.heightEmu > 0) {
                appendLine("        <a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"${pic.widthEmu}\" cy=\"${pic.heightEmu}\"/></a:xfrm>")
            }
            appendLine("        <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>")
            appendLine("      </xdr:spPr>")
            appendLine("    </xdr:pic>")
            appendLine("    <xdr:clientData/>")
            appendLine("  </xdr:twoCellAnchor>")
        }

        appendLine("</xdr:wsDr>")
    }

    private fun buildSheetDrawingRel(drawingIdx: Int): String = buildString {
        appendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        appendLine("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""")
        appendLine("""  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing$drawingIdx.xml"/>""")
        appendLine("</Relationships>")
    }
}
