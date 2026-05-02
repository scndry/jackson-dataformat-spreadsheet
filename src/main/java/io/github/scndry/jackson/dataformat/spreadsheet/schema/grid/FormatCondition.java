package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

/**
 * Incomplete conditional format starting point that requires {@code .style(name)}
 * to bind a {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder}
 * style and produce a {@link ConditionalFormatRule}.
 *
 * <p>Immutable value object. Safe to reuse a single instance across multiple
 * {@link GridConfigurer#conditionalFormatting(String, ConditionalFormatRule, ConditionalFormatRule...)}
 * calls — each {@code .style()} call returns a new {@link ConditionalFormatRule}, the
 * receiver is not mutated. Thread-safe by construction.
 *
 * <p>Subclass extension is package-private (Java 8 has no {@code sealed}). Library users
 * obtain instances via {@link ConditionalFormats} factory methods.
 *
 * @see ConditionalFormats
 * @see ConditionalFormatRule
 * @since 1.6.0
 */
public abstract class FormatCondition {

    FormatCondition() {}

    /**
     * Bind a style name (declared in {@code StylesBuilder}) to produce a complete rule.
     *
     * @param name style name to apply to matching cells
     * @return a new {@link ConditionalFormatRule}; the receiver is unchanged
     */
    public abstract ConditionalFormatRule style(final String name);
}
