package io.github.scndry.jackson.dataformat.spreadsheet.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import lombok.EqualsAndHashCode;

import com.fasterxml.jackson.annotation.JacksonAnnotationValue;

/**
 * Annotates a nested object field to render its flattened columns
 * under a shared group header. The {@link #value()} text becomes
 * a horizontally merged header cell spanning all child columns of
 * the nested object. Annotations stack across nesting depth to form
 * multi-row headers (parent group on top, leaf column names at the
 * bottom).
 *
 * <p>Applies to nested object fields. Collections of nested objects
 * ({@code List<NestedType>}, arrays) are also supported — the group
 * header spans every flattened column produced by the element type.
 * On leaf fields the annotation is ignored (a warning is logged at
 * schema build time).
 *
 * <p>If combined with {@link com.fasterxml.jackson.annotation.JsonUnwrapped @JsonUnwrapped}
 * on the same field, the group is silently ignored.
 *
 * @see DataGrid
 * @see DataColumn
 * @see io.github.scndry.jackson.dataformat.spreadsheet.schema.Column
 */
@Incubating
@Target({ElementType.ANNOTATION_TYPE, ElementType.FIELD, ElementType.METHOD, ElementType.PARAMETER})
@Retention(RetentionPolicy.RUNTIME)
public @interface DataColumnGroup {

    /** Column group header name. */
    String value() default "";

    /** Comment text for the header cell of this column group. */
    String comment() default "";

    @EqualsAndHashCode
    final class Value implements JacksonAnnotationValue<DataColumnGroup> {

        private static final Value EMPTY = new Value("", "");

        private final String _name;
        private final String _comment;

        public Value(final String name, final String comment) {
            _name = name;
            _comment = comment;
        }

        private Value(final DataColumnGroup ann) {
            this(ann.value(), ann.comment());
        }

        public static Value empty() { return EMPTY; }

        public static Value from(final DataColumnGroup ann) {
            return ann == null ? EMPTY : new Value(ann);
        }

        public String getName() { return _name; }

        public String getComment() { return _comment; }

        public boolean isEmpty() { return _name.isEmpty(); }

        @Override
        public Class<DataColumnGroup> valueFor() { return DataColumnGroup.class; }

        @Override
        public String toString() {
            return "DataColumnGroup.Value(name=" + _name + ", comment=" + _comment + ")";
        }
    }

    /**
     * Ordered hierarchy of {@link Value} entries representing the
     * {@code @DataColumnGroup} annotations stacked from the outermost
     * nested-object field down to the leaf column. An empty hierarchy means
     * the column is flat (no enclosing group).
     */
    @EqualsAndHashCode
    final class Hierarchy implements Iterable<Value> {

        private static final Hierarchy EMPTY = new Hierarchy(Collections.emptyList());

        private final List<Value> _groups;

        private Hierarchy(final List<Value> groups) {
            _groups = groups;
        }

        /** Returns the canonical empty hierarchy (no enclosing groups). */
        public static Hierarchy empty() { return EMPTY; }

        /** Number of groups in this hierarchy. {@code 0} for {@link #empty()}. */
        public int depth() { return _groups.size(); }

        public boolean isEmpty() { return _groups.isEmpty(); }

        /**
         * Returns the {@link Value} at the given hierarchy position.
         *
         * @throws IndexOutOfBoundsException if {@code depth < 0 || depth >= depth()}
         */
        public Value at(final int depth) { return _groups.get(depth); }

        /**
         * Returns a sub-hierarchy containing the first {@code depth} groups.
         * {@code parentPath(0)} returns {@link #empty()}; {@code parentPath(depth())}
         * returns this same hierarchy.
         *
         * @throws IndexOutOfBoundsException if {@code depth < 0 || depth > depth()}
         */
        public Hierarchy parentPath(final int depth) {
            if (depth == 0) return EMPTY;
            if (depth == _groups.size()) return this;
            return new Hierarchy(Collections.unmodifiableList(_groups.subList(0, depth)));
        }

        /**
         * Returns a new hierarchy with {@code group} appended at the bottom.
         * If {@code group} is {@link Value#isEmpty() empty}, returns this same
         * instance unchanged (no-op).
         */
        public Hierarchy append(final Value group) {
            if (group.isEmpty()) return this;
            final List<Value> next = new ArrayList<>(_groups.size() + 1);
            next.addAll(_groups);
            next.add(group);
            return new Hierarchy(Collections.unmodifiableList(next));
        }

        @Override
        public Iterator<Value> iterator() { return _groups.iterator(); }

        @Override
        public String toString() {
            return "DataColumnGroup.Hierarchy" + _groups;
        }
    }
}
