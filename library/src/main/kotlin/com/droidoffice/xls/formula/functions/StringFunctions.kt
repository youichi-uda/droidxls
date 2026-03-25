package com.droidoffice.xls.formula.functions

import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.ErrorCode
import com.droidoffice.xls.formula.EvalContext
import com.droidoffice.xls.formula.FormulaFunction
import com.droidoffice.xls.formula.cellValueToString
import com.droidoffice.xls.formula.toNumber

object ConcatenateFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        return CellValue.Text(args.joinToString("") { cellValueToString(it) })
    }
}

object LeftFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val text = cellValueToString(args[0])
        val count = if (args.size > 1) toNumber(args[1]).toInt() else 1
        return CellValue.Text(text.take(count))
    }
}

object RightFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val text = cellValueToString(args[0])
        val count = if (args.size > 1) toNumber(args[1]).toInt() else 1
        return CellValue.Text(text.takeLast(count))
    }
}

object MidFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 3) return CellValue.Error(ErrorCode.VALUE)
        val text = cellValueToString(args[0])
        val start = toNumber(args[1]).toInt() - 1 // Excel MID is 1-based
        val count = toNumber(args[2]).toInt()
        if (start < 0 || count < 0) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Text(text.drop(start).take(count))
    }
}

object LenFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Number(cellValueToString(args[0]).length.toDouble())
    }
}

object TrimFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Text(cellValueToString(args[0]).trim().replace(Regex("\\s+"), " "))
    }
}

object SubstituteFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 3) return CellValue.Error(ErrorCode.VALUE)
        val text = cellValueToString(args[0])
        val oldText = cellValueToString(args[1])
        val newText = cellValueToString(args[2])
        return CellValue.Text(text.replace(oldText, newText))
    }
}

object UpperFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Text(cellValueToString(args[0]).uppercase())
    }
}

object LowerFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Text(cellValueToString(args[0]).lowercase())
    }
}
