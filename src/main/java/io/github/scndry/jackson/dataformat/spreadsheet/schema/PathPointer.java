package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import lombok.EqualsAndHashCode;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

@EqualsAndHashCode
final class PathPointer implements ColumnPointer {

    static final ColumnPointer EMPTY = new PathPointer("");
    static final ColumnPointer ARRAY = new PathPointer("[]");
    static final ColumnPointer PARENT = new PathPointer("..");

    private final Path _path;

    private PathPointer(final String name) {
        _path = Paths.get(name);

    }

    private PathPointer(final Path path) {
        _path = path;
    }

    @Override
    public ColumnPointer resolve(final String other) {
        return new PathPointer(_path.resolve(other));
    }

    @Override
    public ColumnPointer resolve(final ColumnPointer other) {
        return new PathPointer(_path.resolve(path(other)));
    }

    @Override
    public ColumnPointer resolveArray() {
        return resolve(ARRAY);
    }

    @Override
    public ColumnPointer relativize(final ColumnPointer other) {
        return new PathPointer(_path.relativize(path(other)));
    }

    @Override
    public int depth() {
        return _path.getNameCount();
    }

    @Override
    public boolean isEmpty() {
        return EMPTY.equals(this);
    }

    @Override
    public boolean isParent() {
        return PARENT.equals(this);
    }

    @Override
    public ColumnPointer getParent() {
        final Path parent = _path.getParent();
        if (parent == null) return EMPTY;
        return new PathPointer(parent);
    }

    @Override
    public ColumnPointer head() {
        return new PathPointer(_path.getName(0));
    }

    @Override
    public String name() {
        return _path.getFileName().toString();
    }

    @Override
    public boolean contains(final ColumnPointer other) {
        return stream().anyMatch(other::equals);
    }

    @Override
    public Stream<ColumnPointer> stream() {
        if (isEmpty()) {
            return Stream.empty();
        }
        return StreamSupport.stream(_path.spliterator(), false).map(PathPointer::new);
    }

    @Override
    public Iterator<ColumnPointer> iterator() {
        return stream().iterator();
    }

    @Override
    public String toString() {
        return _path.toString();
    }

    @Override
    public boolean startsWith(final ColumnPointer other) {
        return _path.startsWith(path(other));
    }

    private Path path(final ColumnPointer other) {
        return ((PathPointer) other)._path;
    }
}
