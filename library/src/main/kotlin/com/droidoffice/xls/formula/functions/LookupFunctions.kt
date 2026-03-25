package com.droidoffice.xls.formula.functions

import com.droidoffice.xls.core.*
import com.droidoffice.xls.formula.EvalContext
import com.droidoffice.xls.formula.FormulaFunction
import com.droidoffice.xls.formula.cellValueToString
import com.droidoffice.xls.formula.toNumber

object VLookupFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        // VLOOKUP args are parsed as flat values from the range expansion
        // For proper VLOOKUP we need: lookupValue, then table range values, colIndex, [rangeLookup]
        // Simplified: VLOOKUP(value, range_values..., col_index, [range_lookup])
        // This is a simplified implementation
        return CellValue.Error(ErrorCode.NA) // TODO: full range-aware VLOOKUP
    }
}

object HLookupFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        return CellValue.Error(ErrorCode.NA) // TODO: full range-aware HLOOKUP
    }
}

object IndexFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 2) return CellValue.Error(ErrorCode.VALUE)
        val rowIdx = toNumber(args.last()).toInt() - 1
        val data = args.dropLast(1)
        return data.getOrElse(rowIdx) { CellValue.Error(ErrorCode.REF) }
    }
}

object MatchFunction : FormulaFunction {
    override fun execute(args: List<CellValue>, ctx: EvalContext): CellValue {
        if (args.size < 2) return CellValue.Error(ErrorCode.VALUE)
        val lookupValue = args[0]
        val data = args.drop(1)
        val lookupStr = cellValueToString(lookupValue)
        for ((i, v) in data.withIndex()) {
            if (cellValueToString(v) == lookupStr) {
                return CellValue.Number((i + 1).toDouble())
            }
        }
        return CellValue.Error(ErrorCode.NA)
    }
}
