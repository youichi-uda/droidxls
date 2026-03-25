package com.droidoffice.xls.chart

import com.droidoffice.core.drawingml.OfficeColor

/**
 * Represents a chart embedded in a worksheet.
 */
data class Chart(
    val type: ChartType,
    val title: String? = null,
    val series: List<ChartSeries> = emptyList(),
    val fromCol: Int = 0,
    val fromRow: Int = 0,
    val toCol: Int = 10,
    val toRow: Int = 15,
)

enum class ChartType {
    BAR, COLUMN, LINE, PIE, AREA, SCATTER,
}

/**
 * A data series in a chart.
 */
data class ChartSeries(
    val name: String? = null,
    val categoryRange: String? = null,
    val valueRange: String,
    val color: OfficeColor? = null,
)

/**
 * DSL builder for Chart.
 */
class ChartBuilder(private val type: ChartType) {
    var title: String? = null
    var fromCol: Int = 0
    var fromRow: Int = 0
    var toCol: Int = 10
    var toRow: Int = 15
    private val seriesList = mutableListOf<ChartSeries>()

    fun series(valueRange: String, block: ChartSeriesBuilder.() -> Unit = {}) {
        val builder = ChartSeriesBuilder(valueRange)
        builder.apply(block)
        seriesList.add(builder.build())
    }

    internal fun build(): Chart = Chart(
        type = type, title = title,
        series = seriesList.toList(),
        fromCol = fromCol, fromRow = fromRow,
        toCol = toCol, toRow = toRow,
    )
}

class ChartSeriesBuilder(private val valueRange: String) {
    var name: String? = null
    var categoryRange: String? = null
    var color: OfficeColor? = null

    internal fun build(): ChartSeries = ChartSeries(
        name = name, categoryRange = categoryRange,
        valueRange = valueRange, color = color,
    )
}
