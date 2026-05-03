package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.Incubating;

/**
 * Complete conditional formatting rule, accepted by
 * {@link GridConfigurer#conditionalFormatting(String, ConditionalFormatRule, ConditionalFormatRule...)}.
 *
 * <p>Immutable value object. Safe to reuse across multiple
 * {@code conditionalFormatting()} calls or to apply the same rule to multiple
 * columns. Thread-safe by construction.
 *
 * <p>Subclass extension is package-private (Java 8 has no {@code sealed}). Library users
 * obtain instances by completing a {@link FormatCondition} with {@code .style(name)},
 * or directly via visualization factories on {@link ConditionalFormats}
 * (e.g., {@link ConditionalFormats#colorScale()}).
 *
 * @see ConditionalFormats
 * @see FormatCondition
 * @since 1.6.0
 */
@Incubating
public abstract class ConditionalFormatRule {

    ConditionalFormatRule() {}
}
