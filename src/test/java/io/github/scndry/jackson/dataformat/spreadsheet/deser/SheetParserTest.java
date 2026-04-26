package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import com.fasterxml.jackson.core.JsonToken;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.fasterxml.jackson.core.JsonLocation;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetLocation;
import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamReadException;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

import static org.assertj.core.api.Assertions.*;

class SheetParserTest {

    @TempDir File tempDir;
    SpreadsheetMapper mapper;
    SheetParser parser;

    @DataGrid
    static class Flat {
        public int a;
        public int b;
        public Flat() {}
    }

    @DataGrid
    static class Nested {
        public int a;
        public Inner inner;
        public Nested() {}
    }

    static class Inner {
        public int b;
        public Inner() {}
    }

    @BeforeEach
    void setUp() throws Exception {
        mapper = new SpreadsheetMapper();
        File file = createFixture();
        parser = mapper.createParser(file);
    }

    private File createFixture() throws IOException {
        File file = new File(tempDir, "entries.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Entries");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("a");
            header.createCell(1).setCellValue("b");
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue(1);
            r1.createCell(1).setCellValue(2);
            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue(3);
            r2.createCell(1).setCellValue(4);
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
        return file;
    }

    @Test
    void noSchemaThrows() {
        assertThatThrownBy(parser::nextToken)
                .isInstanceOf(SheetStreamReadException.class)
                .hasMessageContaining(SpreadsheetSchema.SCHEMA_TYPE);
    }

    @Test
    void parserCurrentLocationIsSheetLocation() throws Exception {
        parser.setSchema(mapper.sheetSchemaFor(Flat.class));
        // Advance into first row
        parser.nextToken(); // START_ARRAY
        parser.nextToken(); // START_OBJECT
        parser.nextToken(); // FIELD_NAME
        parser.nextToken(); // VALUE_NUMBER_INT

        JsonLocation loc = parser.getCurrentLocation();
        assertThat(loc).isInstanceOf(SheetLocation.class);
        SheetLocation sheetLoc = (SheetLocation) loc;
        assertThat(sheetLoc.getRow()).isEqualTo(1);
        assertThat(sheetLoc.getColumn()).isGreaterThanOrEqualTo(0);
    }

    @Test
    void flatTokenSequence() throws Exception {
        parser.setSchema(mapper.sheetSchemaFor(Flat.class));

        assertThat(parser.hasCurrentToken()).isFalse();
        assertToken(JsonToken.START_ARRAY);

        // row 1
        assertToken(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertToken(JsonToken.END_OBJECT);

        // row 2
        assertToken(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertToken(JsonToken.END_OBJECT);

        assertToken(JsonToken.END_ARRAY);
        assertThat(parser.nextToken()).isNull();

        parser.close();
        assertThat(parser.isClosed()).isTrue();
    }

    @Test
    void nestedTokenSequence() throws Exception {
        parser.setSchema(mapper.sheetSchemaFor(Nested.class));

        assertToken(JsonToken.START_ARRAY);

        // row 1: { a, inner: { b } }
        assertToken(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertField(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertToken(JsonToken.END_OBJECT);
        assertToken(JsonToken.END_OBJECT);

        // row 2
        assertToken(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertField(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertToken(JsonToken.END_OBJECT);
        assertToken(JsonToken.END_OBJECT);

        assertToken(JsonToken.END_ARRAY);
        assertThat(parser.nextToken()).isNull();

        parser.close();
    }

    private void assertField(JsonToken valueToken) throws IOException {
        assertToken(JsonToken.FIELD_NAME);
        assertToken(valueToken);
    }

    private void assertToken(JsonToken expected) throws IOException {
        assertThat(parser.nextToken()).isEqualTo(expected);
        assertThat(parser.hasCurrentToken()).isTrue();
    }
}
