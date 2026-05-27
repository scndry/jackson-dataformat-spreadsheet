package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.SpreadsheetVersion;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Schema-build-time validation for the {@code shift} attribute on
 * {@code @DataColumn} and {@code @DataColumnGroup}: rejects negative
 * values and values exceeding {@link SpreadsheetVersion#EXCEL2007} max
 * columns. All violations are collected and reported together.
 *
 * <p>Not part of the public API. Classes under
 * {@code io.github.scndry.jackson.dataformat.spreadsheet.schema.internal}
 * may change without notice between releases — do not invoke directly
 * from application code.
 */
public final class SchemaShiftValidator {

    private static final int MAX_SHIFT = SpreadsheetVersion.EXCEL2007.getMaxColumns();

    private SchemaShiftValidator() {
    }

    /** Throws {@link IllegalStateException} listing every violation when the
     *  schema breaks the shift contract. */
    public static void validate(final SpreadsheetSchema schema) {
        final List<String> violations = new ArrayList<>();
        final Set<DataColumnGroup.Value> seenGroups = new HashSet<>();
        for (final Column column : schema) {
            if (column == null) continue;
            _check(violations, "Column '" + column.getName() + "'", column.getValue().getShift());
            for (final DataColumnGroup.Value group : column.getGroupHierarchy()) {
                if (seenGroups.add(group)) {
                    _check(violations, "Column group '" + group.getName() + "'", group.getShift());
                }
            }
        }
        if (violations.isEmpty()) return;
        final StringBuilder sb = new StringBuilder("Shift invariant violation:");
        for (final String v : violations) sb.append("\n - ").append(v);
        sb.append("\n\nshift must be in [0, ").append(MAX_SHIFT).append("].");
        throw new IllegalStateException(sb.toString());
    }

    private static void _check(final List<String> out, final String subject, final int shift) {
        if (shift < 0) {
            out.add(subject + " has invalid shift " + shift + " (must be >= 0)");
        } else if (shift > MAX_SHIFT) {
            out.add(subject + " has invalid shift " + shift + " (must be <= " + MAX_SHIFT + ")");
        }
    }
}
