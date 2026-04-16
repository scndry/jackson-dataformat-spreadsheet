package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.Iterator;
import java.util.stream.Stream;

/**
 * Hierarchical path that identifies a column's position within
 * a nested JSON structure. Supports resolution, relativization,
 * and iteration over path segments, similar to
 * {@link java.nio.file.Path} for file systems.
 *
 * @see Column
 * @see SpreadsheetSchema
 */
public interface ColumnPointer extends Iterable<ColumnPointer> {

    static ColumnPointer empty() {
        return SegmentPointer.EMPTY;
    }

    static ColumnPointer array() {
        return SegmentPointer.ARRAY;
    }

    ColumnPointer resolve(String other);

    ColumnPointer resolve(ColumnPointer other);

    ColumnPointer resolveArray();

    ColumnPointer relativize(ColumnPointer other);

    int depth();

    boolean isEmpty();

    boolean isParent();

    ColumnPointer getParent();

    ColumnPointer head();

    String name();

    boolean contains(ColumnPointer other);

    Stream<ColumnPointer> stream();

    @Override
    Iterator<ColumnPointer> iterator();

    @Override
    String toString();

    boolean startsWith(ColumnPointer other);

    boolean equals(Object o);

    int hashCode();
}
