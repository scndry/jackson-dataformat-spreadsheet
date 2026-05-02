package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import java.util.Objects;

/**
 * Internal expression {@link FormatCondition}: boolean Excel formula awaiting
 * {@code .style(name)} to produce an {@link ExpressionRule}.
 */
final class ExpressionCondition extends FormatCondition {

    final String _formula;

    ExpressionCondition(final String formula) {
        _formula = formula;
    }

    @Override
    public ConditionalFormatRule style(final String name) {
        Objects.requireNonNull(name, "style name");
        return new ExpressionRule(_formula, name);
    }
}
