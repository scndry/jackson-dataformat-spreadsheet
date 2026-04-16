package io.github.scndry.jackson.dataformat.spreadsheet.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import lombok.EqualsAndHashCode;

import com.fasterxml.jackson.annotation.JacksonAnnotationValue;
import com.fasterxml.jackson.annotation.OptBoolean;

/**
 * Marks a type as the root of a spreadsheet data grid. Required
 * on any class used as the root type for
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema}
 * generation. Attributes define default column styling, width
 * constraints, and merge behavior for all columns in the grid.
 *
 * @see DataColumn
 * @see io.github.scndry.jackson.dataformat.spreadsheet.SchemaGenerator
 */
@Target(
        {ElementType.ANNOTATION_TYPE,
        ElementType.TYPE,
        ElementType.METHOD,
        ElementType.CONSTRUCTOR,
        ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface DataGrid {

    int DEFAULT_COLUMN_WIDTH = -1;
    int DEFAULT_MIN_COLUMN_WIDTH = -1;
    int DEFAULT_MAX_COLUMN_WIDTH = 255;

    /** Default cell style name for data cells. */
    String columnStyle() default "";

    /** Default cell style name for header cells. */
    String columnHeaderStyle() default "";

    /** Default column width in character units ({@code -1} = auto). */
    int columnWidth() default DEFAULT_COLUMN_WIDTH;

    /** Whether to auto-size columns to fit content. */
    OptBoolean autoSizeColumn() default OptBoolean.DEFAULT;

    /** Minimum column width ({@code -1} = no minimum). */
    int minColumnWidth() default DEFAULT_MIN_COLUMN_WIDTH;

    /** Maximum column width (default 255). */
    int maxColumnWidth() default DEFAULT_MAX_COLUMN_WIDTH;

    /** Whether to merge cells vertically for repeated values. */
    OptBoolean mergeColumn() default OptBoolean.DEFAULT;

    @EqualsAndHashCode
    final class Value implements JacksonAnnotationValue<DataGrid> {

        private static final Value EMPTY = new Value();

        private final String _columnStyle;
        private final String _columnHeaderStyle;
        private final int _columnWidth;
        private final OptBoolean _autoSizeColumn;
        private final int _minColumnWidth;
        private final int _maxColumnWidth;
        private final OptBoolean _mergeColumn;

        public Value(final String columnStyle, final String columnHeaderStyle,
                     final int columnWidth, final OptBoolean autoSizeColumn,
                     final int minColumnWidth, final int maxColumnWidth,
                     final OptBoolean mergeColumn) {
            _columnStyle = columnStyle;
            _columnHeaderStyle = columnHeaderStyle;
            _columnWidth = columnWidth;
            _autoSizeColumn = autoSizeColumn;
            _minColumnWidth = minColumnWidth;
            _maxColumnWidth = maxColumnWidth;
            _mergeColumn = mergeColumn;
        }

        private Value() {
            this("", "", DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT,
                    DEFAULT_MIN_COLUMN_WIDTH, DEFAULT_MAX_COLUMN_WIDTH, OptBoolean.DEFAULT);
        }

        private Value(final DataGrid ann) {
            this(ann.columnStyle(), ann.columnHeaderStyle(), ann.columnWidth(),
                    ann.autoSizeColumn(), ann.minColumnWidth(), ann.maxColumnWidth(),
                    ann.mergeColumn());
        }

        public static Value empty() { return EMPTY; }

        public static Value from(final DataGrid ann) {
            return ann == null ? EMPTY : new Value(ann);
        }

        public String getColumnStyle() { return _columnStyle; }
        public String getColumnHeaderStyle() { return _columnHeaderStyle; }
        public int getColumnWidth() { return _columnWidth; }
        public OptBoolean getAutoSizeColumn() { return _autoSizeColumn; }
        public int getMinColumnWidth() { return _minColumnWidth; }
        public int getMaxColumnWidth() { return _maxColumnWidth; }
        public OptBoolean getMergeColumn() { return _mergeColumn; }

        public boolean isEmpty() {
            return EMPTY.equals(this);
        }

        public Value withDefaults(final Value defaults) {
            if (defaults.isEmpty() || defaults.equals(this)) return this;
            return new Value(
                    _columnStyle.isEmpty() ? defaults._columnStyle : _columnStyle,
                    _columnHeaderStyle.isEmpty() ? defaults._columnHeaderStyle : _columnHeaderStyle,
                    _columnWidth == DEFAULT_COLUMN_WIDTH ? defaults._columnWidth : _columnWidth,
                    _autoSizeColumn == OptBoolean.DEFAULT
                            ? defaults._autoSizeColumn : _autoSizeColumn,
                    _minColumnWidth == DEFAULT_MIN_COLUMN_WIDTH
                            ? defaults._minColumnWidth : _minColumnWidth,
                    _maxColumnWidth == DEFAULT_MAX_COLUMN_WIDTH
                            ? defaults._maxColumnWidth : _maxColumnWidth,
                    _mergeColumn == OptBoolean.DEFAULT ? defaults._mergeColumn : _mergeColumn
            );
        }

        @Override
        public Class<DataGrid> valueFor() { return DataGrid.class; }

        @Override
        public String toString() {
            return "DataGrid.Value(columnStyle=" + _columnStyle
                    + ", columnHeaderStyle=" + _columnHeaderStyle
                    + ", columnWidth=" + _columnWidth
                    + ", autoSizeColumn=" + _autoSizeColumn
                    + ", minColumnWidth=" + _minColumnWidth
                    + ", maxColumnWidth=" + _maxColumnWidth
                    + ", mergeColumn=" + _mergeColumn + ")";
        }
    }
}
