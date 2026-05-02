package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

/**
 * Operand reference value type for conditional formatting cellIs rules. Produced via
 * {@link ConditionalFormats#formula(String)} (raw Excel formula passthrough) or
 * {@link ConditionalFormats#columnRef(String)} (schema-aware row-relative reference).
 *
 * <p>Immutable value object. Thread-safe by construction.
 *
 * @see ConditionalFormats#formula(String)
 * @see ConditionalFormats#columnRef(String)
 * @since 1.6.0
 */
public final class Formula {

    enum Kind { OF, COLUMN }

    private final Kind _kind;
    private final String _value;

    Formula(final Kind kind, final String value) {
        _kind = kind;
        _value = value;
    }

    Kind kind() {
        return _kind;
    }

    String value() {
        return _value;
    }
}
