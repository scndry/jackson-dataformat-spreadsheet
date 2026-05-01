package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

/**
 * Operand reference for conditional formatting rules.
 * Use {@link #of(String)} for raw Excel formula passthrough (cell references, functions, expressions).
 * Use {@link #column(String)} for schema-aware row-relative references to another schema column.
 *
 * @see ConditionalFormattingRuleSpec
 */
public final class Formula {

    enum Kind { OF, COLUMN }

    private final Kind _kind;
    private final String _value;

    private Formula(final Kind kind, final String value) {
        _kind = kind;
        _value = value;
    }

    /**
     * Raw Excel formula passthrough. The value is emitted verbatim into the OOXML
     * {@code <formula>} element. The library does not validate or parse it.
     *
     * <pre>{@code
     * Formula.of("$D$1")
     * Formula.of("AVERAGE($B$2:$B$100)")
     * Formula.of("$E1 * 0.9")
     * }</pre>
     */
    public static Formula of(final String text) {
        if (text == null || text.isEmpty()) {
            throw new IllegalArgumentException("Formula text must not be empty");
        }
        return new Formula(Kind.OF, text);
    }

    /**
     * Schema-aware row-relative reference to another column in the data grid.
     * Resolved to {@code $<colLetter><dataStartRow>} at apply time, so Excel
     * auto-shifts the row for each cell in the conditional formatting range.
     *
     * <pre>{@code
     * .column("price").greaterThan(Formula.column("minPrice"))   // this row's price > this row's minPrice
     * }</pre>
     */
    public static Formula column(final String name) {
        if (name == null || name.isEmpty()) {
            throw new IllegalArgumentException("Column name must not be empty");
        }
        return new Formula(Kind.COLUMN, name);
    }

    Kind kind() {
        return _kind;
    }

    String value() {
        return _value;
    }
}
