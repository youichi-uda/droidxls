package com.droidoffice.xls.formula.functions

import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.ErrorCode
import com.droidoffice.xls.formula.EvalContext
import com.droidoffice.xls.formula.FormulaFunction
import com.droidoffice.xls.formula.toNumber
import kotlin.math.*

object RoundFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val num = toNumber(args[0])
        val digits = if (args.size > 1) toNumber(args[1]).toInt() else 0
        val factor = 10.0.pow(digits)
        return CellValue.Number(Math.round(num * factor) / factor)
    }
}

object RoundUpFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val num = toNumber(args[0])
        val digits = if (args.size > 1) toNumber(args[1]).toInt() else 0
        val factor = 10.0.pow(digits)
        val rounded = if (num >= 0) ceil(num * factor) / factor else floor(num * factor) / factor
        return CellValue.Number(rounded)
    }
}

object RoundDownFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val num = toNumber(args[0])
        val digits = if (args.size > 1) toNumber(args[1]).toInt() else 0
        val factor = 10.0.pow(digits)
        val rounded = if (num >= 0) floor(num * factor) / factor else ceil(num * factor) / factor
        return CellValue.Number(rounded)
    }
}

object AbsFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Number(abs(toNumber(args[0])))
    }
}

object ModFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 2) return CellValue.Error(ErrorCode.VALUE)
        val divisor = toNumber(args[1])
        if (divisor == 0.0) return CellValue.Error(ErrorCode.DIV_ZERO)
        return CellValue.Number(toNumber(args[0]) % divisor)
    }
}

object IntFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Number(floor(toNumber(args[0])))
    }
}

object CeilingFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val num = toNumber(args[0])
        val sig = if (args.size > 1) toNumber(args[1]) else 1.0
        if (sig == 0.0) return CellValue.Number(0.0)
        return CellValue.Number(ceil(num / sig) * sig)
    }
}

object FloorFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val num = toNumber(args[0])
        val sig = if (args.size > 1) toNumber(args[1]) else 1.0
        if (sig == 0.0) return CellValue.Error(ErrorCode.DIV_ZERO)
        return CellValue.Number(floor(num / sig) * sig)
    }
}
