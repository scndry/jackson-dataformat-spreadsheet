package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import org.apache.poi.ss.SpreadsheetVersion;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Schema-build-time validation for the {@code shift} attribute on
 * {@code @DataColumn} and {@code @DataColumnGroup}: rejects negative
 * values and values exceeding {@link SpreadsheetVersion#EXCEL2007} max
 * columns.
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

    /** Throws {@link IllegalStateException} on the first invalid shift. */
    public static void validate(final SpreadsheetSchema schema) {
        for (final Column column : schema) {
            if (column == null) continue;
            _check("Column '" + column.getName() + "'", column.getValue().getShift());
            for (final DataColumnGroup.Value group : column.getGroupHierarchy()) {
                _check("Column group '" + group.getName() + "'", group.getShift());
            }
        }
    }

    private static void _check(final String subject, final int shift) {
        if (shift < 0) {
            throw new IllegalStateException(
                    subject + " has invalid shift " + shift + " (must be >= 0)");
        }
        if (shift > MAX_SHIFT) {
            throw new IllegalStateException(
                    subject + " has invalid shift " + shift + " (must be <= " + MAX_SHIFT + ")");
        }
    }
}
