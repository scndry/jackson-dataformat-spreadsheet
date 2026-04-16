package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.Arrays;
import java.util.Iterator;
import java.util.stream.Stream;

/**
 * String array-based {@link ColumnPointer} implementation. Lazy-caches parent, head, and element
 * pointers for reduced allocation in hot paths.
 */
final class SegmentPointer implements ColumnPointer {

    private static final String[] EMPTY_SEGMENTS = new String[0];

    static final ColumnPointer EMPTY = new SegmentPointer(EMPTY_SEGMENTS);
    static final ColumnPointer ARRAY = new SegmentPointer(new String[]{"[]"});
    static final ColumnPointer PARENT = new SegmentPointer(new String[]{".."});

    private final String[] _segments;

    // Lazy cached — derived from immutable _segments, safe like String.hashCode
    private ColumnPointer _parent;
    private ColumnPointer _head;
    private ColumnPointer[] _elements;

    private SegmentPointer(final String[] segments) {
        _segments = segments;
    }

    @Override
    public ColumnPointer resolve(final String other) {
        if (_segments.length == 0) return new SegmentPointer(new String[]{other});
        final String[] result = Arrays.copyOf(_segments, _segments.length + 1);
        result[_segments.length] = other;
        return new SegmentPointer(result);
    }

    @Override
    public ColumnPointer resolve(final ColumnPointer other) {
        final String[] otherSegs = segments(other);
        if (_segments.length == 0) return other;
        if (otherSegs.length == 0) return this;
        final String[] result = new String[_segments.length + otherSegs.length];
        System.arraycopy(_segments, 0, result, 0, _segments.length);
        System.arraycopy(otherSegs, 0, result, _segments.length, otherSegs.length);
        return new SegmentPointer(result);
    }

    @Override
    public ColumnPointer resolveArray() {
        return resolve(ARRAY);
    }

    @Override
    public ColumnPointer relativize(final ColumnPointer other) {
        final String[] b = segments(other);
        if (_segments.length == 0) return other;
        final String[] a = _segments;
        int common = 0;
        while (common < a.length && common < b.length && a[common].equals(b[common])) {
            common++;
        }
        final int ups = a.length - common;
        final int downs = b.length - common;
        if (ups == 0 && downs == 0) return EMPTY;
        final String[] result = new String[ups + downs];
        Arrays.fill(result, 0, ups, "..");
        System.arraycopy(b, common, result, ups, downs);
        return new SegmentPointer(result);
    }

    @Override
    public int depth() {
        return _segments.length == 0 ? 1 : _segments.length;
    }

    @Override
    public boolean isEmpty() {
        return _segments.length == 0;
    }

    @Override
    public boolean isParent() {
        return _segments.length == 1 && "..".equals(_segments[0]);
    }

    @Override
    public ColumnPointer getParent() {
        ColumnPointer p = _parent;
        if (p == null) {
            p = _segments.length <= 1
                    ? EMPTY : new SegmentPointer(Arrays.copyOf(_segments, _segments.length - 1));
            _parent = p;
        }
        return p;
    }

    @Override
    public ColumnPointer head() {
        ColumnPointer h = _head;
        if (h == null) {
            h = new SegmentPointer(new String[]{_segments[0]});
            _head = h;
        }
        return h;
    }

    @Override
    public String name() {
        return _segments[_segments.length - 1];
    }

    @Override
    public boolean contains(final ColumnPointer other) {
        return stream().anyMatch(other::equals);
    }

    @Override
    public Stream<ColumnPointer> stream() {
        return Arrays.stream(_elements());
    }

    @Override
    public Iterator<ColumnPointer> iterator() {
        return Arrays.asList(_elements()).iterator();
    }

    private ColumnPointer[] _elements() {
        ColumnPointer[] e = _elements;
        if (e == null) {
            e = new ColumnPointer[_segments.length];
            for (int i = 0; i < _segments.length; i++) {
                e[i] = new SegmentPointer(new String[]{_segments[i]});
            }
            _elements = e;
        }
        return e;
    }

    @Override
    public boolean startsWith(final ColumnPointer other) {
        final String[] otherSegs = segments(other);
        if (otherSegs.length == 0) return true;
        if (otherSegs.length > _segments.length) return false;
        for (int i = 0; i < otherSegs.length; i++) {
            if (!_segments[i].equals(otherSegs[i])) return false;
        }
        return true;
    }

    @Override
    public String toString() {
        return String.join("/", _segments);
    }

    @Override
    public boolean equals(final Object o) {
        if (this == o) return true;
        if (!(o instanceof SegmentPointer)) return false;
        return Arrays.equals(_segments, ((SegmentPointer) o)._segments);
    }

    @Override
    public int hashCode() {
        return Arrays.hashCode(_segments);
    }

    private static String[] segments(final ColumnPointer other) {
        return ((SegmentPointer) other)._segments;
    }
}
