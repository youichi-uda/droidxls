package com.droidoffice.xls.formula

import com.droidoffice.xls.core.*
import com.droidoffice.xls.formula.functions.*

/**
 * Evaluates cell formulas. Supports ~30 built-in functions.
 */
class FormulaEvaluator(private val worksheet: Worksheet) {

    private val functions = buildMap<String, FormulaFunction> {
        // Aggregate
        put("SUM", SumFunction); put("AVERAGE", AverageFunction)
        put("COUNT", CountFunction); put("COUNTA", CountAFunction)
        put("COUNTBLANK", CountBlankFunction); put("MAX", MaxFunction); put("MIN", MinFunction)
        // Logic
        put("IF", IfFunction); put("AND", AndFunction); put("OR", OrFunction)
        put("NOT", NotFunction); put("IFERROR", IfErrorFunction); put("IFNA", IfNaFunction)
        // Math
        put("ROUND", RoundFunction); put("ROUNDUP", RoundUpFunction); put("ROUNDDOWN", RoundDownFunction)
        put("ABS", AbsFunction); put("MOD", ModFunction); put("INT", IntFunction)
        put("CEILING", CeilingFunction); put("FLOOR", FloorFunction)
        // String
        put("CONCATENATE", ConcatenateFunction); put("LEFT", LeftFunction); put("RIGHT", RightFunction)
        put("MID", MidFunction); put("LEN", LenFunction); put("TRIM", TrimFunction)
        put("SUBSTITUTE", SubstituteFunction); put("UPPER", UpperFunction); put("LOWER", LowerFunction)
        // Lookup
        put("VLOOKUP", VLookupFunction); put("HLOOKUP", HLookupFunction)
        put("INDEX", IndexFunction); put("MATCH", MatchFunction)
        // Date
        put("TODAY", TodayFunction); put("NOW", NowFunction); put("DATE", DateFunction)
        put("YEAR", YearFunction); put("MONTH", MonthFunction); put("DAY", DayFunction)
    }

    /**
     * Evaluate a cell's formula and return the result as a CellValue.
     */
    fun evaluate(cell: Cell): CellValue {
        val formula = cell.cellValue as? CellValue.Formula ?: return cell.cellValue
        return try {
            val tokens = FormulaTokenizer.tokenize(formula.expression)
            val ctx = EvalContext(worksheet, this, functions)
            val result = ExpressionEvaluator.evaluate(tokens, ctx)
            result
        } catch (_: Exception) {
            CellValue.Error(ErrorCode.VALUE)
        }
    }

    /**
     * Recalculate all formulas in the worksheet.
     */
    fun recalculateAll() {
        for (cell in worksheet.cells()) {
            if (cell.cellValue is CellValue.Formula) {
                val result = evaluate(cell)
                val formula = cell.cellValue as CellValue.Formula
                cell.cellValue = formula.copy(cachedValue = result)
            }
        }
    }
}

data class EvalContext(
    val worksheet: Worksheet,
    val evaluator: FormulaEvaluator,
    val functions: Map<String, FormulaFunction>,
)

interface FormulaFunction {
    fun execute(args: List<CellValue>, ctx: EvalContext): CellValue
}

/**
 * Simple recursive descent expression evaluator.
 */
object ExpressionEvaluator {

    fun evaluate(tokens: List<Token>, ctx: EvalContext): CellValue {
        val parser = Parser(tokens, ctx)
        return parser.parseExpression()
    }

    class Parser(private val tokens: List<Token>, private val ctx: EvalContext) {
        private var pos = 0

        private fun peek(): Token? = tokens.getOrNull(pos)
        private fun advance(): Token = tokens[pos++]

        fun parseExpression(): CellValue = parseComparison()

        private fun parseComparison(): CellValue {
            var left = parseAddSub()
            while (peek() is Token.Operator && (peek() as Token.Operator).op in listOf("=", "<", ">", "<=", ">=", "<>", "!=")) {
                val op = (advance() as Token.Operator).op
                val right = parseAddSub()
                left = compareValues(left, right, op)
            }
            return left
        }

        private fun parseAddSub(): CellValue {
            var left = parseMulDiv()
            while (peek() is Token.Operator && (peek() as Token.Operator).op in listOf("+", "-", "&")) {
                val op = (advance() as Token.Operator).op
                val right = parseMulDiv()
                left = if (op == "&") {
                    CellValue.Text(cellValueToString(left) + cellValueToString(right))
                } else {
                    binaryMath(left, right, op)
                }
            }
            return left
        }

        private fun parseMulDiv(): CellValue {
            var left = parsePower()
            while (peek() is Token.Operator && (peek() as Token.Operator).op in listOf("*", "/")) {
                val op = (advance() as Token.Operator).op
                val right = parsePower()
                left = binaryMath(left, right, op)
            }
            return left
        }

        private fun parsePower(): CellValue {
            var left = parseAtom()
            while (peek() is Token.Operator && (peek() as Token.Operator).op == "^") {
                advance()
                val right = parseAtom()
                left = CellValue.Number(Math.pow(toNumber(left), toNumber(right)))
            }
            return left
        }

        fun parseAtom(): CellValue {
            return when (val t = peek()) {
                is Token.Number -> { advance(); CellValue.Number(t.value) }
                is Token.StringLiteral -> { advance(); CellValue.Text(t.value) }
                is Token.BooleanLiteral -> { advance(); CellValue.Bool(t.value) }
                is Token.CellRef -> { advance(); resolveCellRef(t.ref) }
                is Token.Range -> { advance(); CellValue.Empty } // ranges only meaningful in function args
                is Token.Function -> { advance(); parseFunction(t.name) }
                is Token.LeftParen -> {
                    advance()
                    val result = parseExpression()
                    if (peek() is Token.RightParen) advance()
                    result
                }
                else -> { if (pos < tokens.size) advance(); CellValue.Empty }
            }
        }

        private fun parseFunction(name: String): CellValue {
            // Expect (
            if (peek() is Token.LeftParen) advance()

            val args = mutableListOf<CellValue>()
            while (peek() != null && peek() !is Token.RightParen) {
                // Check if arg is a range
                val t = peek()
                if (t is Token.Range) {
                    advance()
                    val rangeValues = resolveRange(t.from, t.to)
                    args.addAll(rangeValues)
                } else {
                    args.add(parseExpression())
                }
                if (peek() is Token.Comma) advance()
            }
            if (peek() is Token.RightParen) advance()

            val func = ctx.functions[name]
                ?: return CellValue.Error(ErrorCode.NAME)
            return func.execute(args, ctx)
        }

        private fun resolveCellRef(ref: String): CellValue {
            val cellRef = try { CellReference.parse(ref) } catch (_: Exception) { return CellValue.Empty }
            val cell = ctx.worksheet.getCellOrNull(cellRef) ?: return CellValue.Empty
            // Recursively evaluate formulas
            if (cell.cellValue is CellValue.Formula) {
                return ctx.evaluator.evaluate(cell)
            }
            return cell.cellValue
        }

        private fun resolveRange(from: String, to: String): List<CellValue> {
            val fromRef = try { CellReference.parse(from) } catch (_: Exception) { return emptyList() }
            val toRef = try { CellReference.parse(to) } catch (_: Exception) { return emptyList() }
            val values = mutableListOf<CellValue>()
            for (r in fromRef.row..toRef.row) {
                for (c in fromRef.col..toRef.col) {
                    val cell = ctx.worksheet.getCellOrNull(CellReference(r, c))
                    if (cell != null) {
                        val v = if (cell.cellValue is CellValue.Formula) ctx.evaluator.evaluate(cell) else cell.cellValue
                        values.add(v)
                    } else {
                        values.add(CellValue.Empty)
                    }
                }
            }
            return values
        }

        private fun binaryMath(left: CellValue, right: CellValue, op: String): CellValue {
            val l = toNumber(left)
            val r = toNumber(right)
            return CellValue.Number(when (op) {
                "+" -> l + r
                "-" -> l - r
                "*" -> l * r
                "/" -> if (r == 0.0) return CellValue.Error(ErrorCode.DIV_ZERO) else l / r
                else -> 0.0
            })
        }

        private fun compareValues(left: CellValue, right: CellValue, op: String): CellValue {
            val cmp = when {
                left is CellValue.Number && right is CellValue.Number -> left.value.compareTo(right.value)
                left is CellValue.Text && right is CellValue.Text -> left.value.compareTo(right.value, ignoreCase = true)
                else -> toNumber(left).compareTo(toNumber(right))
            }
            return CellValue.Bool(when (op) {
                "=" -> cmp == 0
                "<>" , "!=" -> cmp != 0
                "<" -> cmp < 0
                ">" -> cmp > 0
                "<=" -> cmp <= 0
                ">=" -> cmp >= 0
                else -> false
            })
        }
    }
}

fun toNumber(value: CellValue): Double = when (value) {
    is CellValue.Number -> value.value
    is CellValue.Bool -> if (value.value) 1.0 else 0.0
    is CellValue.Text -> value.value.toDoubleOrNull() ?: 0.0
    is CellValue.DateValue -> DateUtil.dateTimeToSerial(value.value)
    else -> 0.0
}

fun cellValueToString(value: CellValue): String = when (value) {
    is CellValue.Text -> value.value
    is CellValue.Number -> if (value.value == value.value.toLong().toDouble()) value.value.toLong().toString() else value.value.toString()
    is CellValue.Bool -> value.value.toString().uppercase()
    is CellValue.Empty -> ""
    is CellValue.Error -> value.code.symbol
    is CellValue.DateValue -> value.value.toString()
    is CellValue.Formula -> value.cachedValue?.let { cellValueToString(it) } ?: ""
}

fun toBool(value: CellValue): Boolean = when (value) {
    is CellValue.Bool -> value.value
    is CellValue.Number -> value.value != 0.0
    is CellValue.Text -> value.value.isNotEmpty()
    else -> false
}
