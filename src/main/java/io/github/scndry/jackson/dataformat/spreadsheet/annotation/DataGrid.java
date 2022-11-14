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

@Target({ElementType.ANNOTATION_TYPE, ElementType.TYPE, ElementType.METHOD, ElementType.CONSTRUCTOR, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface DataGrid {

    int DEFAULT_COLUMN_WIDTH = -1;
    int DEFAULT_MIN_COLUMN_WIDTH = -1;
    int DEFAULT_MAX_COLUMN_WIDTH = 255;

    String columnStyle() default "";

    String columnHeaderStyle() default "";

    int columnWidth() default DEFAULT_COLUMN_WIDTH;

    OptBoolean autoSizeColumn() default OptBoolean.DEFAULT;

    int minColumnWidth() default DEFAULT_MIN_COLUMN_WIDTH;

    int maxColumnWidth() default DEFAULT_MAX_COLUMN_WIDTH;

    @Getter
    @ToString
    @EqualsAndHashCode
    @RequiredArgsConstructor
    final class Value implements JacksonAnnotationValue<DataGrid> {

        private static final Value EMPTY = new Value();

        private final String columnStyle;
        private final String columnHeaderStyle;
        private final int columnWidth;
        private final OptBoolean autoSizeColumn;
        private final int minColumnWidth;
        private final int maxColumnWidth;

        private Value() {
            this("", "", DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT, DEFAULT_MIN_COLUMN_WIDTH, DEFAULT_MAX_COLUMN_WIDTH);
        }

        private Value(final DataGrid ann) {
            this(ann.columnStyle(), ann.columnHeaderStyle(), ann.columnWidth(), ann.autoSizeColumn(), ann.minColumnWidth(), ann.maxColumnWidth());
        }

        public static Value empty() {
            return EMPTY;
        }

        public static Value from(final DataGrid ann) {
            return ann == null ? EMPTY : new Value(ann);
        }

        public boolean isEmpty() {
            return EMPTY.equals(this);
        }

        public Value withDefaults(final Value defaults) {
            if (defaults.isEmpty() || defaults.equals(this)) return this;
            return new Value(
                    columnStyle.isEmpty() ? defaults.columnStyle : columnStyle,
                    columnHeaderStyle.isEmpty() ? defaults.columnHeaderStyle : columnHeaderStyle,
                    columnWidth == DataGrid.DEFAULT_COLUMN_WIDTH ? defaults.columnWidth : columnWidth,
                    autoSizeColumn == OptBoolean.DEFAULT ? defaults.autoSizeColumn : autoSizeColumn,
                    minColumnWidth == DataGrid.DEFAULT_MIN_COLUMN_WIDTH ? defaults.minColumnWidth : minColumnWidth,
                    maxColumnWidth == DataGrid.DEFAULT_MAX_COLUMN_WIDTH ? defaults.maxColumnWidth : maxColumnWidth
            );
        }

        @Override
        public Class<DataGrid> valueFor() {
            return DataGrid.class;
        }
    }
}
