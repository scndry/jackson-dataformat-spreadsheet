package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Lookup of POI {@link CellStyle} instances by registered name or by Java type.
 * Built per-workbook by
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder}.
 *
 * @see SpreadsheetSchema
 */
public interface Styles {

    /** Returns the style registered for {@code name}, or {@code null} if none. */
    CellStyle byName(String name);

    /** Returns the style registered for {@code type}, or {@code null} if none. */
    CellStyle byType(Class<?> type);

    /** Resolves {@code name} via {@link #byName}, falling back to {@link #byType}
     *  with {@code typeFallback} when {@code name} is empty. Throws
     *  {@link NoSuchElementException} when {@code name} is non-empty and unregistered. */
    default CellStyle resolve(final String name, final Class<?> typeFallback) {
        if (!name.isEmpty()) {
            final CellStyle cs = byName(name);
            if (cs == null) {
                throw new NoSuchElementException("Style '" + name + "' is not registered");
            }
            return cs;
        }
        return typeFallback == null ? null : byType(typeFallback);
    }
}
