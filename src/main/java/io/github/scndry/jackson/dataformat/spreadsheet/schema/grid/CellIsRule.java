package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

/**
 * Internal cellIs ({@code CELL_VALUE_IS}) rule: comparison operator + 1 or 2 operands + style name.
 *
 * <p>Operands are stored as either pre-formatted Excel formula {@code String}s (for
 * value-typed factories like {@code greaterThan(80)}) or {@link Formula} objects (for
 * raw / column-relative references). Resolution to an Excel formula happens at apply
 * time in {@link GridConfigurer}.
 */
final class CellIsRule extends ConditionalFormatRule {

    final byte _operator;
    final Object _operand1;
    final Object _operand2;
    final String _styleName;

    CellIsRule(final byte operator, final Object operand1, final Object operand2, final String styleName) {
        _operator = operator;
        _operand1 = operand1;
        _operand2 = operand2;
        _styleName = styleName;
    }
}
