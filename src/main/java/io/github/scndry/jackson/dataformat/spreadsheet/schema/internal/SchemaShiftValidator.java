package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Schema-build-time validation for the {@code shift} attribute on
 * {@code @DataColumn} and {@code @DataColumnGroup}.
 *
 * <p>Rule: {@code shift < 0} is rejected.
 *
 * <p>Not part of the public API. Classes under
 * {@code io.github.scndry.jackson.dataformat.spreadsheet.schema.internal}
 * may change without notice between releases — do not invoke directly
 * from application code.
 */
public final class SchemaShiftValidator {

    private SchemaShiftValidator() {
    }

    /** Throws {@link IllegalStateException} when the schema violates the shift contract. */
    public static void validate(final SpreadsheetSchema schema) {
        for (final Column column : schema) {
            if (column == null) continue;
            final int columnShift = column.getValue().getShift();
            if (columnShift < 0) {
                throw new IllegalStateException(
                        "Column '" + column.getName() + "' has invalid shift "
                                + columnShift + " (must be >= 0).");
            }
            for (final DataColumnGroup.Value group : column.getGroupHierarchy()) {
                final int groupShift = group.getShift();
                if (groupShift < 0) {
                    throw new IllegalStateException(
                            "Column group '" + group.getName() + "' has invalid shift "
                                    + groupShift + " (must be >= 0).");
                }
            }
        }
    }
}
