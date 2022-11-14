package io.github.scndry.jackson.dataformat.spreadsheet.annotation;

import com.fasterxml.jackson.annotation.JacksonAnnotationValue;
import com.fasterxml.jackson.annotation.OptBoolean;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.ToString;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.ANNOTATION_TYPE, ElementType.FIELD, ElementType.METHOD, ElementType.PARAMETER})
@Retention(RetentionPolicy.RUNTIME)
public @interface DataColumn {

    /**
     * Defines column name
     */
    String value() default "";

    String style() default "";

    String headerStyle() default "";

    int width() default DataGrid.DEFAULT_COLUMN_WIDTH;

    OptBoolean autoSize() default OptBoolean.DEFAULT;

    int minWidth() default DataGrid.DEFAULT_MIN_COLUMN_WIDTH;

    int maxWidth() default DataGrid.DEFAULT_MAX_COLUMN_WIDTH;

    @Getter
    @ToString
    @EqualsAndHashCode
    @RequiredArgsConstructor
    final class Value implements JacksonAnnotationValue<DataColumn> {

        private static final Value EMPTY = new Value();

        private final String name;
        private final String style;
        private final String headerStyle;
        private final int width;
        private final OptBoolean autoSize;
        private final int minWidth;
        private final int maxWidth;

        private Value() {
            this("", "", "", DataGrid.DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT, DataGrid.DEFAULT_MIN_COLUMN_WIDTH, DataGrid.DEFAULT_MAX_COLUMN_WIDTH);
        }

        private Value(final DataColumn ann) {
            this(ann.value(), ann.style(), ann.headerStyle(), ann.width(), ann.autoSize(), ann.minWidth(), ann.maxWidth());
        }

        public static Value empty() {
            return EMPTY;
        }

        public static Value from(final DataColumn ann) {
            return ann == null ? EMPTY : new Value(ann);
        }

        public Value withName(final String name) {
            if (name == null || name.isEmpty()) return this;
            return new Value(name, style, headerStyle, width, autoSize, minWidth, maxWidth);
        }

        public Value withDefaults(final DataGrid.Value defaults) {
            if (defaults.isEmpty()) return this;
            return new Value(name,
                    style.isEmpty() ? defaults.getColumnStyle() : style,
                    headerStyle.isEmpty() ? defaults.getColumnHeaderStyle() : headerStyle,
                    width == DataGrid.DEFAULT_COLUMN_WIDTH ? defaults.getColumnWidth() : width,
                    autoSize == OptBoolean.DEFAULT ? defaults.getAutoSizeColumn() : autoSize,
                    minWidth == DataGrid.DEFAULT_MIN_COLUMN_WIDTH ? defaults.getMinColumnWidth() : minWidth,
                    maxWidth == DataGrid.DEFAULT_MAX_COLUMN_WIDTH ? defaults.getMaxColumnWidth() : maxWidth
            );
        }

        public boolean isAutoSize() {
            return width < 0 && autoSize == OptBoolean.TRUE;
        }

        @Override
        public Class<DataColumn> valueFor() {
            return DataColumn.class;
        }
    }
}
