package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import java.util.Objects;

/**
 * Internal cellIs {@link FormatCondition}: comparison operator + 1 or 2 operands,
 * awaiting {@code .style(name)} to produce a {@link CellIsRule}.
 */
final class CellIsCondition extends FormatCondition {

    final byte _operator;
    final Object _operand1;
    final Object _operand2;

    CellIsCondition(final byte operator, final Object operand1, final Object operand2) {
        _operator = operator;
        _operand1 = operand1;
        _operand2 = operand2;
    }

    @Override
    public ConditionalFormatRule style(final String name) {
        Objects.requireNonNull(name, "style name");
        return new CellIsRule(_operator, _operand1, _operand2, name);
    }
}
