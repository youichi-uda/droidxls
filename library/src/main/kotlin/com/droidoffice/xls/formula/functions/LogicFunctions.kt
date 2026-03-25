package com.droidoffice.xls.formula.functions

import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.ErrorCode
import com.droidoffice.xls.formula.EvalContext
import com.droidoffice.xls.formula.FormulaFunction
import com.droidoffice.xls.formula.toBool

object IfFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        val condition = toBool(args[0])
        return if (condition) args.getOrElse(1) { CellValue.Bool(true) }
        else args.getOrElse(2) { CellValue.Bool(false) }
    }
}

object AndFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        return CellValue.Bool(args.all { toBool(it) })
    }
}

object OrFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        return CellValue.Bool(args.any { toBool(it) })
    }
}

object NotFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.isEmpty()) return CellValue.Error(ErrorCode.VALUE)
        return CellValue.Bool(!toBool(args[0]))
    }
}

object IfErrorFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 2) return CellValue.Error(ErrorCode.VALUE)
        return if (args[0] is CellValue.Error) args[1] else args[0]
    }
}

object IfNaFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 2) return CellValue.Error(ErrorCode.VALUE)
        val isNa = args[0] is CellValue.Error && (args[0] as CellValue.Error).code == ErrorCode.NA
        return if (isNa) args[1] else args[0]
    }
}
