package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;

import com.fasterxml.jackson.databind.exc.InvalidDefinitionException;

import lombok.Data;
import lombok.Getter;
import lombok.RequiredArgsConstructor;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.AnnotatedNameResolver;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.ColumnNameResolver;

import static org.assertj.core.api.Assertions.*;

class SchemaGeneratorTest {

    @TempDir File tempDir;
    File out;
    Row row;

    @DataGrid
    static class SimpleEntry {
        public int a;
        public int b;
        public SimpleEntry() {}
        public SimpleEntry(int a, int b) {
            this.a = a; this.b = b;
        }
    }

    @BeforeEach
    void setUp() {
        out = new File(tempDir, "test.xlsx");
    }

    @Test
    void originAddress() throws Exception {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .origin("B2").build();
        mapper.writeValue(out, new SimpleEntry(1, 2));
        try (XSSFWorkbook wb = new XSSFWorkbook(out)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat((Object) sheet.getRow(0)).isNull();
            row = sheet.getRow(1);
            assertThat(row.getCell(0)).isNull();
            assertCell(1, "a");
            assertCell(2, "b");
            row = sheet.getRow(2);
            assertThat(row.getCell(0)).isNull();
            assertCell(1, 1);
            assertCell(2, 2);
        }
    }

    @Test
    void columnNameResolver() throws Exception {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .columnNameResolver(
                        prop -> prop.getName().toUpperCase())
                .build();
        mapper.writeValue(out, new SimpleEntry(1, 2));
        try (XSSFWorkbook wb = new XSSFWorkbook(out)) {
            row = wb.getSheetAt(0).getRow(0);
            assertCell(0, "A");
            assertCell(1, "B");
        }
    }

    @Test
    void annotatedColumnNames() throws Exception {
        ColumnNameResolver resolver =
                ColumnNameResolver.annotated(
                        NameOf.class, ColumnCode::getText);
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .columnNameResolver(resolver).build();
        mapper.writeValue(out, null, AnnotatedEntity.class);
        try (XSSFWorkbook wb = new XSSFWorkbook(out)) {
            row = wb.getSheetAt(0).getRow(0);
            assertCell(0, ColumnCode.A.getText());
            assertCell(1, ColumnCode.B.getText());
        }
    }

    @Test
    void missingAnnotationThrows() {
        ColumnNameResolver resolver =
                AnnotatedNameResolver.forValue(
                        NameOf.class, ColumnCode::getText);
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .columnNameResolver(resolver).build();
        assertThatThrownBy(() -> mapper.writeValue(
                out, null, MissingAnnotationEntity.class))
                .isInstanceOf(InvalidDefinitionException.class)
                .hasMessageContaining("@NameOf");
    }

    private void assertCell(int col, String expected) {
        assertThat(row.getCell(col).getStringCellValue())
                .isEqualTo(expected);
    }

    private void assertCell(int col, double expected) {
        assertThat(row.getCell(col).getNumericCellValue())
                .isEqualTo(expected);
    }

    // -- test types --

    @Getter
    @RequiredArgsConstructor
    enum ColumnCode {
        A("Code A"), B("Code B");
        private final String text;
    }

    @java.lang.annotation.Retention(
            java.lang.annotation.RetentionPolicy.RUNTIME)
    @interface NameOf {
        ColumnCode value();
    }

    @Data @DataGrid
    class AnnotatedEntity {
        @NameOf(ColumnCode.A)
        @DataColumn("overwritten") int a;
        @NameOf(ColumnCode.B) int b;
    }

    @Data @DataGrid
    class MissingAnnotationEntity {
        int a;
        @NameOf(ColumnCode.B) int b;
    }
}
