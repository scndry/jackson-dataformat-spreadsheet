package io.github.scndry.jackson.dataformat.spreadsheet.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import lombok.EqualsAndHashCode;

import com.fasterxml.jackson.annotation.JacksonAnnotationValue;
import com.fasterxml.jackson.annotation.OptBoolean;

/**
 * Annotates a field or property to customize its spreadsheet
 * column representation. Allows overriding the column name,
 * style, width, and merge behavior on a per-column basis.
 * Unset attributes inherit defaults from the enclosing
 * {@link DataGrid}.
 *
 * @see DataGrid
 * @see io.github.scndry.jackson.dataformat.spreadsheet.schema.Column
 */
@Target({ElementType.ANNOTATION_TYPE, ElementType.FIELD, ElementType.METHOD, ElementType.PARAMETER})
@Retention(RetentionPolicy.RUNTIME)
public @interface DataColumn {

    /**
     * Defines column name
     */
    String value() default "";

    /** Cell style name for data cells in this column. */
    String style() default "";

    /** Cell style name for the header cell of this column. */
    String headerStyle() default "";

    /** Column width in character units ({@code -1} = auto). */
    int width() default DataGrid.DEFAULT_COLUMN_WIDTH;

    /** Whether to auto-size this column to fit content. */
    OptBoolean autoSize() default OptBoolean.DEFAULT;

    /** Minimum column width ({@code -1} = no minimum). */
    int minWidth() default DataGrid.DEFAULT_MIN_COLUMN_WIDTH;

    /** Maximum column width (default 255). */
    int maxWidth() default DataGrid.DEFAULT_MAX_COLUMN_WIDTH;

    /** Whether to merge cells vertically for repeated values. */
    OptBoolean merge() default OptBoolean.DEFAULT;

    /** Comment text for the header cell of this column. */
    String comment() default "";

    @EqualsAndHashCode
    final class Value implements JacksonAnnotationValue<DataColumn> {

        private static final Value EMPTY = new Value();

        private final String _name;
        private final String _style;
        private final String _headerStyle;
        private final int _width;
        private final OptBoolean _autoSize;
        private final int _minWidth;
        private final int _maxWidth;
        private final OptBoolean _merge;
        private final String _comment;

        public Value(final String name, final String style, final String headerStyle,
                     final int width, final OptBoolean autoSize,
                     final int minWidth, final int maxWidth, final OptBoolean merge,
                     final String comment) {
            _name = name;
            _style = style;
            _headerStyle = headerStyle;
            _width = width;
            _autoSize = autoSize;
            _minWidth = minWidth;
            _maxWidth = maxWidth;
            _merge = merge;
            _comment = comment;
        }

        private Value() {
            this("", "", "", DataGrid.DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT,
                    DataGrid.DEFAULT_MIN_COLUMN_WIDTH, DataGrid.DEFAULT_MAX_COLUMN_WIDTH,
                    OptBoolean.DEFAULT, "");
        }

        private Value(final DataColumn ann) {
            this(ann.value(), ann.style(), ann.headerStyle(), ann.width(),
                    ann.autoSize(), ann.minWidth(), ann.maxWidth(), ann.merge(),
                    ann.comment());
        }

        public static Value empty() { return EMPTY; }

        public static Value from(final DataColumn ann) {
            return ann == null ? EMPTY : new Value(ann);
        }

        public String getName() { return _name; }
        public String getStyle() { return _style; }
        public String getHeaderStyle() { return _headerStyle; }
        public int getWidth() { return _width; }
        public OptBoolean getAutoSize() { return _autoSize; }
        public int getMinWidth() { return _minWidth; }
        public int getMaxWidth() { return _maxWidth; }
        public OptBoolean getMerge() { return _merge; }
        public String getComment() { return _comment; }

        public Value withName(final String name) {
            if (name == null || name.isEmpty()) return this;
            return new Value(name, _style, _headerStyle, _width, _autoSize,
                    _minWidth, _maxWidth, _merge, _comment);
        }

        public Value withDefaults(final DataGrid.Value defaults) {
            if (defaults.isEmpty()) return this;
            return new Value(_name,
                    _style.isEmpty() ? defaults.getColumnStyle() : _style,
                    _headerStyle.isEmpty() ? defaults.getColumnHeaderStyle() : _headerStyle,
                    _width == DataGrid.DEFAULT_COLUMN_WIDTH ? defaults.getColumnWidth() : _width,
                    _autoSize == OptBoolean.DEFAULT ? defaults.getAutoSizeColumn() : _autoSize,
                    _minWidth == DataGrid.DEFAULT_MIN_COLUMN_WIDTH
                            ? defaults.getMinColumnWidth() : _minWidth,
                    _maxWidth == DataGrid.DEFAULT_MAX_COLUMN_WIDTH
                            ? defaults.getMaxColumnWidth() : _maxWidth,
                    _merge == OptBoolean.DEFAULT ? defaults.getMergeColumn() : _merge,
                    _comment
            );
        }

        public boolean isAutoSize() {
            return _width < 0 && _autoSize == OptBoolean.TRUE;
        }

        public boolean isMerge() {
            return _merge == OptBoolean.TRUE;
        }

        @Override
        public Class<DataColumn> valueFor() { return DataColumn.class; }

        @Override
        public String toString() {
            return "DataColumn.Value(name=" + _name
                    + ", style=" + _style
                    + ", headerStyle=" + _headerStyle
                    + ", width=" + _width
                    + ", autoSize=" + _autoSize
                    + ", minWidth=" + _minWidth
                    + ", maxWidth=" + _maxWidth
                    + ", merge=" + _merge
                    + ", comment=" + _comment + ")";
        }
    }
}
