package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.List;

import com.fasterxml.jackson.core.FormatSchema;

/**
 * {@link FormatSchema} for spreadsheet reading and writing. Produced by
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper#sheetSchemaFor}
 * and passed back into the mapper's reader / writer chains. Iterates the
 * ordered {@link Column} layout (sibling-format pattern: see
 * {@code CsvSchema}).
 *
 * @see Column
 * @see io.github.scndry.jackson.dataformat.spreadsheet.SchemaGenerator
 */
public interface SpreadsheetSchema extends FormatSchema, Iterable<Column> {

    String SCHEMA_TYPE = "spreadsheet";

    int FEATURE_USE_HEADER = 0x0001;
    int FEATURE_COLUMN_REORDERING = 0x0002;

    int DEFAULT_FEATURES = FEATURE_USE_HEADER;

    int size();

    Column column(int index);

    Column column(String name);

    int columnIndex(String name);

    String columnName(int index);

    List<String> columnNames();
}
