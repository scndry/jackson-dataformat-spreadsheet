package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Read-time validation for the depth-aware nested-list contract:
 * each nested-list-bearing record level must declare exactly one
 * {@code @DataColumn(anchor = true)} column at its immediate scope.
 *
 * <p>Not part of the public API. Classes under
 * {@code io.github.scndry.jackson.dataformat.spreadsheet.schema.internal}
 * may change without notice between releases — do not invoke directly
 * from application code.
 */
public final class NestedAnchorValidator {

    private NestedAnchorValidator() {
    }

    /** Throws {@link IllegalStateException} when the schema violates the
     *  anchor invariant — exactly one anchor per nested-list-bearing
     *  record level, no anchors elsewhere. */
    public static void validate(final SpreadsheetSchema schema) {
        final Set<ColumnPointer> expected = _listBearingScopes(schema);
        final Map<ColumnPointer, List<Column>> anchorsByScope = _anchorsByScope(schema);

        final List<ColumnPointer> missing = new ArrayList<>();
        for (final ColumnPointer scope : expected) {
            if (!anchorsByScope.containsKey(scope)) missing.add(scope);
        }
        final List<ColumnPointer> extra = new ArrayList<>();
        final List<Map.Entry<ColumnPointer, List<Column>>> duplicate = new ArrayList<>();
        for (final Map.Entry<ColumnPointer, List<Column>> e : anchorsByScope.entrySet()) {
            if (!expected.contains(e.getKey())) extra.add(e.getKey());
            else if (e.getValue().size() > 1) duplicate.add(e);
        }
        if (missing.isEmpty() && extra.isEmpty() && duplicate.isEmpty()) return;
        throw new IllegalStateException(_buildMessage(missing, extra, duplicate));
    }

    private static Set<ColumnPointer> _listBearingScopes(final SpreadsheetSchema schema) {
        final Set<ColumnPointer> scopes = new LinkedHashSet<>();
        for (final Column c : schema) {
            if (c == null) continue;
            ColumnPointer parent = ColumnPointer.empty();
            ColumnPointer head = ColumnPointer.empty();
            for (final ColumnPointer seg : c.getPointer()) {
                head = head.resolve(seg);
                if (seg.equals(ColumnPointer.array())) {
                    scopes.add(parent);
                    parent = head;
                }
            }
        }
        return scopes;
    }

    private static Map<ColumnPointer, List<Column>> _anchorsByScope(final SpreadsheetSchema schema) {
        final Map<ColumnPointer, List<Column>> result = new LinkedHashMap<>();
        for (final Column c : schema) {
            if (c == null || !c.isAnchor()) continue;
            result.computeIfAbsent(SchemaAnchorInspector.immediateScope(c.getPointer()),
                    k -> new ArrayList<>()).add(c);
        }
        return result;
    }

    private static String _buildMessage(
            final List<ColumnPointer> missing,
            final List<ColumnPointer> extra,
            final List<Map.Entry<ColumnPointer, List<Column>>> duplicate) {
        final StringBuilder sb = new StringBuilder("Anchor invariant violation:");
        for (final ColumnPointer scope : missing) {
            sb.append("\n - Missing anchor at scope '").append(_scopeName(scope))
                    .append("': record contains a nested list");
        }
        for (final ColumnPointer scope : extra) {
            sb.append("\n - Extra anchor at scope '").append(_scopeName(scope))
                    .append("': record has no nested list");
        }
        for (final Map.Entry<ColumnPointer, List<Column>> e : duplicate) {
            final List<String> names = new ArrayList<>();
            for (final Column c : e.getValue()) names.add(c.getName());
            sb.append("\n - Multiple anchors at scope '").append(_scopeName(e.getKey()))
                    .append("': ").append(String.join(", ", names));
        }
        sb.append("\n\nEach nested-list-bearing record level requires exactly one"
                + " @DataColumn(anchor = true) column at its immediate scope.");
        return sb.toString();
    }

    private static String _scopeName(final ColumnPointer scope) {
        final String s = scope.toString();
        return s.isEmpty() ? "<root>" : s;
    }
}
