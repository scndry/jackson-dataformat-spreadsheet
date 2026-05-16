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

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;

/**
 * Registry of named {@link CellStyleBuilder} instances that builds a {@link Styles} collection from a {@link Workbook}.
 */
public final class StylesBuilder implements Builder<Styles> {

    // Each name/type maps to one POI CellStyle — bounded by Excel's cell-style cap (4K HSSF / 64K XSSF).
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
        return _register(type, new CellStyleBuilder());
    }

    public CellStyleBuilder cellStyle(final String name) {
        return _register(name, new CellStyleBuilder());
    }

    public CellStyleBuilder cellStyle(final String name, final String cloneStyleFrom) {
        final CellStyleBuilder source = _builders.get(cloneStyleFrom);
        if (source == null) {
            throw new IllegalArgumentException("Style '" + cloneStyleFrom + "' is not registered");
        }
        return _register(name, source.copy());
    }

    private CellStyleBuilder _register(final Object key, final CellStyleBuilder builder) {
        builder.setBuilder(this);
        _builders.put(key, builder);
        return builder;
    }

    @Override
    public Styles build(final Workbook workbook) {
        final StylesImpl styles = new StylesImpl();
        _builders.forEach((key, builder) -> styles._styles.put(key, builder.build(workbook)));
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

        @Override
        public CellStyle byName(final String name) {
            return _styles.get(name);
        }

        @Override
        public CellStyle byType(final Class<?> type) {
            return _styles.get(type);
        }
    }
}
