package com.droidoffice.xls.formula.functions

import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.DateUtil
import com.droidoffice.xls.core.ErrorCode
import com.droidoffice.xls.formula.EvalContext
import com.droidoffice.xls.formula.FormulaFunction
import com.droidoffice.xls.formula.toNumber
import java.time.LocalDate
import java.time.LocalDateTime

object TodayFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        return CellValue.DateValue(LocalDate.now().atStartOfDay())
    }
}

object NowFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        return CellValue.DateValue(LocalDateTime.now())
    }
}

object DateFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 3) return CellValue.Error(ErrorCode.VALUE)
        val year = toNumber(args[0]).toInt()
        val month = toNumber(args[1]).toInt()
        val day = toNumber(args[2]).toInt()
        return try {
            CellValue.DateValue(LocalDate.of(year, month, day).atStartOfDay())
        } catch (_: Exception) {
            CellValue.Error(ErrorCode.VALUE)
        }
    }
}

object YearFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val dt = when (val v = args[0]) {
            is CellValue.DateValue -> v.value
            is CellValue.Number -> DateUtil.serialToDateTime(v.value)
            else -> return CellValue.Error(ErrorCode.VALUE)
        }
        return CellValue.Number(dt.year.toDouble())
    }
}

object MonthFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val dt = when (val v = args[0]) {
            is CellValue.DateValue -> v.value
            is CellValue.Number -> DateUtil.serialToDateTime(v.value)
            else -> return CellValue.Error(ErrorCode.VALUE)
        }
        return CellValue.Number(dt.monthValue.toDouble())
    }
}

object DayFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val dt = when (val v = args[0]) {
            is CellValue.DateValue -> v.value
            is CellValue.Number -> DateUtil.serialToDateTime(v.value)
            else -> return CellValue.Error(ErrorCode.VALUE)
        }
        return CellValue.Number(dt.dayOfMonth.toDouble())
    }
}
