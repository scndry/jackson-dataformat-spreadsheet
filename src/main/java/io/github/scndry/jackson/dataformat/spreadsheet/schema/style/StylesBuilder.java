package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;

/**
 * Registry of named {@link CellStyleBuilder} instances that builds a {@link Styles} collection from a {@link Workbook}.
 */
public final class StylesBuilder implements Builder<Styles> {

    private final Map<Object, CellStyleBuilder> _builders;

    public StylesBuilder() {
        this(new HashMap<>());
    }

    private StylesBuilder(final Map<Object, CellStyleBuilder> builders) {
        _builders = builders;
    }

    public static StylesBuilder simple() {
        return new StylesBuilder()
                .cellStyle(Float.TYPE).dataFormat().numberFloatWithComma().end()
                .cellStyle(Double.TYPE).dataFormat().numberFloatWithComma().end()
                .cellStyle(Integer.TYPE).dataFormat().numberIntWithComma().end()
                .cellStyle(Long.TYPE).dataFormat().numberIntWithComma().end()
                .cellStyle(Float.class).dataFormat().numberFloat().end()
                .cellStyle(Double.class).dataFormat().numberFloat().end()
                .cellStyle(Integer.class).dataFormat().numberInt().end()
                .cellStyle(Long.class).dataFormat().numberInt().end()
                .cellStyle(BigInteger.class).dataFormat().text().end()
                .cellStyle(BigDecimal.class).dataFormat().text().end()
                .cellStyle(String.class).dataFormat().text().end()
                .cellStyle(Date.class).dataFormat().dateTime().end()
                .cellStyle(Calendar.class).dataFormat().dateTime().end()
                .cellStyle(LocalDateTime.class).dataFormat().dateTime().end()
                .cellStyle(LocalDate.class).dataFormat().date().end();
    }

    public CellStyleBuilder cellStyle(final Class<?> type) {
        return cellStyle(type, new CellStyleBuilder());
    }

    private CellStyleBuilder cellStyle(final Object key, final CellStyleBuilder builder) {
        builder.setBuilder(this);
        _builders.put(key, builder);
        return builder;
    }

    public CellStyleBuilder cellStyle(final String name) {
        return cellStyle(name, new CellStyleBuilder());
    }

    public CellStyleBuilder cellStyle(final String name, final String cloneStyleFrom) {
        final CellStyleBuilder source = _builders.get(cloneStyleFrom);
        if (source == null) {
            throw new IllegalArgumentException("Style '" + cloneStyleFrom + "' is not registered");
        }
        return cellStyle(name, source.copy());
    }

    @Override
    public Styles build(final Workbook workbook) {
        final StylesImpl styles = new StylesImpl();
        _builders.forEach((key, builder) -> styles.register(key, builder.build(workbook)));
        return styles;
    }

    @Override
    public StylesBuilder copy() {
        final Map<Object, CellStyleBuilder> builders = new HashMap<>();
        _builders.forEach((key, builder) -> builders.put(key, builder.copy()));
        return new StylesBuilder(builders);
    }

    private static final class StylesImpl implements Styles {

        private final Map<Object, CellStyle> _styles = new HashMap<>();

        public void register(final Object key, final CellStyle style) {
            _styles.put(key, style);
        }

        @Override
        public CellStyle getStyle(final Column column) {
            return _findStyle(column.getValue().getStyle(), column);
        }

        @Override
        public CellStyle getStyle(final String name) {
            return _styles.get(name);
        }

        @Override
        public CellStyle getHeaderStyle(final Column column) {
            return _findStyle(column.getValue().getHeaderStyle(), column);
        }

        private CellStyle _findStyle(final String name, final Column column) {
            if (!name.isEmpty()) {
                final CellStyle style = _styles.get(name);
                if (style == null) {
                    throw new IllegalStateException("Style '" + name + "' is not registered");
                }
                return style;
            }
            return _styles.get(column.getType().getRawClass());
        }
    }
}
