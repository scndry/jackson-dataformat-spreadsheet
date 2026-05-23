package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Schema queries that support nested-list reads — anchor lookup,
 * array-scope enumeration, and the {@code immediateScope} /
 * {@code allArrayScopes} pointer helpers. Extracted from
 * {@link SpreadsheetSchema} so the schema retains only its core
 * column data plus the hot-path resolve cache; nested-list-specific
 * lookups live next to {@link BackWriteProjection} and
 * {@link NestedAnchorValidator}.
 *
 * <p>Not part of the public API. Classes under
 * {@code io.github.scndry.jackson.dataformat.spreadsheet.schema.internal}
 * may change without notice between releases — do not invoke directly
 * from application code.
 */
public final class SchemaAnchorInspector {

    private SchemaAnchorInspector() {
    }

    public static boolean hasAnchor(final SpreadsheetSchema schema) {
        for (final Column col : schema) {
            if (col != null && col.isAnchor()) return true;
        }
        return false;
    }

    public static boolean hasNestedList(final SpreadsheetSchema schema) {
        for (final Column col : schema) {
            if (col != null && col.getPointer().contains(ColumnPointer.array())) return true;
        }
        return false;
    }

    public static Column findAnchorColumn(final SpreadsheetSchema schema, final ColumnPointer scope) {
        for (final Column col : schema) {
            if (col == null || !col.isAnchor()) continue;
            if (immediateScope(col.getPointer()).equals(scope)) return col;
        }
        return null;
    }

    public static Set<ColumnPointer> allArrayScopes(final SpreadsheetSchema schema) {
        final Set<ColumnPointer> scopes = new LinkedHashSet<>();
        for (final Column col : schema) {
            if (col == null) continue;
            scopes.addAll(allArrayScopes(col.getPointer()));
        }
        return scopes;
    }

    public static ColumnPointer immediateScope(final ColumnPointer pointer) {
        ColumnPointer head = ColumnPointer.empty();
        ColumnPointer last = ColumnPointer.empty();
        for (final ColumnPointer seg : pointer) {
            head = head.resolve(seg);
            if (seg.equals(ColumnPointer.array())) last = head;
        }
        return last;
    }

    public static List<ColumnPointer> allArrayScopes(final ColumnPointer pointer) {
        final List<ColumnPointer> scopes = new ArrayList<>();
        ColumnPointer head = ColumnPointer.empty();
        for (final ColumnPointer seg : pointer) {
            head = head.resolve(seg);
            if (seg.equals(ColumnPointer.array())) scopes.add(head);
        }
        return scopes;
    }
}
