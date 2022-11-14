package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.Iterator;
import java.util.stream.Stream;

public interface ColumnPointer extends Iterable<ColumnPointer> {

    static ColumnPointer empty() {
        return PathPointer.EMPTY;
    }

    static ColumnPointer array() {
        return PathPointer.ARRAY;
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
