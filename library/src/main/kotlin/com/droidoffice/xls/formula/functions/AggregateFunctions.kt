package com.droidoffice.xls.formula.functions

import com.droidoffice.xls.core.CellValue
import com.droidoffice.xls.core.ErrorCode
import com.droidoffice.xls.formula.EvalContext
import com.droidoffice.xls.formula.FormulaFunction
import com.droidoffice.xls.formula.toNumber

object SumFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        var sum = 0.0
        for (arg in args) {
            if (arg is CellValue.Number || arg is CellValue.Bool) sum += toNumber(arg)
        }
        return CellValue.Number(sum)
    }
}

object AverageFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        var sum = 0.0; var count = 0
        for (arg in args) {
            if (arg is CellValue.Number) { sum += arg.value; count++ }
        }
        return if (count > 0) CellValue.Number(sum / count) else CellValue.Error(ErrorCode.DIV_ZERO)
    }
}

object CountFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        val count = args.count { it is CellValue.Number }
        return CellValue.Number(count.toDouble())
    }
}

object CountAFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        val count = args.count { it !is CellValue.Empty }
        return CellValue.Number(count.toDouble())
    }
}

object CountBlankFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        val count = args.count { it is CellValue.Empty }
        return CellValue.Number(count.toDouble())
    }
}

object MaxFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        val nums = args.filterIsInstance<CellValue.Number>().map { it.value }
        return if (nums.isEmpty()) CellValue.Number(0.0) else CellValue.Number(nums.max())
    }
}

object MinFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        val nums = args.filterIsInstance<CellValue.Number>().map { it.value }
        return if (nums.isEmpty()) CellValue.Number(0.0) else CellValue.Number(nums.min())
    }
}
