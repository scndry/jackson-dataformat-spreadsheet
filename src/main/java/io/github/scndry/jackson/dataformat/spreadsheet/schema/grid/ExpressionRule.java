package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

/**
 * Internal expression ({@code FORMULA}) rule: boolean Excel formula + style name.
 */
final class ExpressionRule extends ConditionalFormatRule {

    final String _formula;
    final String _styleName;

    ExpressionRule(final String formula, final String styleName) {
        _formula = formula;
        _styleName = styleName;
    }
}
