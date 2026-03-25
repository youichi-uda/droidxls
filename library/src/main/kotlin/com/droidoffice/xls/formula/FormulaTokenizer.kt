package com.droidoffice.xls.formula

/**
 * Tokenizes a formula string into tokens for the formula parser.
 */
sealed class Token {
    data class Number(val value: Double) : Token()
    data class StringLiteral(val value: String) : Token()
    data class CellRef(val ref: String) : Token()
    data class Range(val from: String, val to: String) : Token()
    data class Function(val name: String) : Token()
    data class BooleanLiteral(val value: Boolean) : Token()
    data object LeftParen : Token()
    data object RightParen : Token()
    data object Comma : Token()
    data class Operator(val op: String) : Token()
}

object FormulaTokenizer {

    fun tokenize(formula: String): List<Token> {
        val tokens = mutableListOf<Token>()
        var i = 0
        val input = formula.trim()

        while (i < input.length) {
            when {
                input[i].isWhitespace() -> i++
                input[i] == '(' -> { tokens.add(Token.LeftParen); i++ }
                input[i] == ')' -> { tokens.add(Token.RightParen); i++ }
                input[i] == ',' -> { tokens.add(Token.Comma); i++ }
                input[i] in "+-" && (tokens.isEmpty() || tokens.last() is Token.LeftParen || tokens.last() is Token.Comma || tokens.last() is Token.Operator) -> {
                    // Unary +/- at start or after operator/paren
                    val start = i
                    i++
                    while (i < input.length && (input[i].isDigit() || input[i] == '.')) i++
                    val numStr = input.substring(start, i)
                    tokens.add(Token.Number(numStr.toDoubleOrNull() ?: 0.0))
                }
                input[i] in "+-*/^&" -> {
                    tokens.add(Token.Operator(input[i].toString())); i++
                }
                input[i] == '=' || input[i] == '<' || input[i] == '>' || input[i] == '!' -> {
                    val start = i
                    i++
                    if (i < input.length && input[i] == '=') i++
                    tokens.add(Token.Operator(input.substring(start, i)))
                }
                input[i] == '"' -> {
                    i++
                    val sb = StringBuilder()
                    while (i < input.length && input[i] != '"') {
                        sb.append(input[i]); i++
                    }
                    if (i < input.length) i++ // skip closing "
                    tokens.add(Token.StringLiteral(sb.toString()))
                }
                input[i].isDigit() || input[i] == '.' -> {
                    val start = i
                    while (i < input.length && (input[i].isDigit() || input[i] == '.' || input[i] == 'E' || input[i] == 'e' ||
                            (i > start && input[i] in "+-" && input[i-1] in "Ee"))) i++
                    tokens.add(Token.Number(input.substring(start, i).toDouble()))
                }
                input[i].isLetter() || input[i] == '_' || input[i] == '$' -> {
                    val start = i
                    while (i < input.length && (input[i].isLetterOrDigit() || input[i] == '_' || input[i] == '$')) i++
                    val word = input.substring(start, i)

                    // Check for TRUE/FALSE
                    if (word.equals("TRUE", ignoreCase = true)) {
                        tokens.add(Token.BooleanLiteral(true))
                    } else if (word.equals("FALSE", ignoreCase = true)) {
                        tokens.add(Token.BooleanLiteral(false))
                    }
                    // Check if followed by ( → function
                    else if (i < input.length && input[i] == '(') {
                        tokens.add(Token.Function(word.uppercase()))
                    }
                    // Check for range (A1:B2)
                    else if (i < input.length && input[i] == ':') {
                        val from = word
                        i++ // skip :
                        val rangeStart = i
                        while (i < input.length && (input[i].isLetterOrDigit() || input[i] == '$')) i++
                        val to = input.substring(rangeStart, i)
                        tokens.add(Token.Range(from.replace("$", ""), to.replace("$", "")))
                    }
                    // Otherwise cell reference
                    else {
                        tokens.add(Token.CellRef(word.replace("$", "")))
                    }
                }
                else -> i++ // skip unknown
            }
        }
        return tokens
    }
}
