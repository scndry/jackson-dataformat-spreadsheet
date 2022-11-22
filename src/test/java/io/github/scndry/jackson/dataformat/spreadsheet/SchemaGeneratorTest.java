package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.databind.exc.InvalidDefinitionException;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.AnnotatedNameResolver;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.ColumnNameResolver;
import lombok.Data;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import support.fixture.Entry;

import java.io.File;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SchemaGeneratorTest {

    @TempDir
    Path tempDir;
    File out;
    Row row;

    @BeforeEach
    void setUp() throws Exception {
        out = tempDir.resolve("test.xlsx").toFile();
    }

    @Test
    void changeOriginAddress() throws Exception {
        final SpreadsheetMapper mapper = new SpreadsheetMapper().setOrigin("B2");
        mapper.writeValue(out, Entry.VALUE);
        try (XSSFWorkbook workbook = new XSSFWorkbook(out)) {
            final XSSFSheet sheet = workbook.getSheetAt(0);
            row = sheet.getRow(0);
            assertThat(row).isNull();
            row = sheet.getRow(1);
            assertThat(row.getCell(0)).isNull();
            assertCellValue(1, "a");
            assertCellValue(2, "b");
            row = sheet.getRow(2);
            assertThat(row.getCell(0)).isNull();
            assertCellValue(1, 1);
            assertCellValue(2, 2);
            row = sheet.getRow(3);
            assertThat(row).isNull();
        }
    }

    @Test
    void overwriteColumnNames() throws Exception {
        final SpreadsheetMapper mapper = new SpreadsheetMapper()
                .setColumnNameResolver(prop -> prop.getName().toUpperCase());
        mapper.writeValue(out, Entry.VALUE);
        try (XSSFWorkbook workbook = new XSSFWorkbook(out)) {
            final XSSFSheet sheet = workbook.getSheetAt(0);
            row = sheet.getRow(0);
            assertCellValue(0, Entry.Fields.A.toUpperCase());
            assertCellValue(1, Entry.Fields.B.toUpperCase());
        }
    }

    @Test
    void annotatedColumnNames() throws Exception {
        ColumnNameResolver byText = AnnotatedNameResolver.forValue(NameOf.class, ColumnCode::getText);
        final SpreadsheetMapper mapper = new SpreadsheetMapper()
                .setColumnNameResolver(byText);
        mapper.writeValue(out, null, AnnotatedEntity.class);
        try (XSSFWorkbook workbook = new XSSFWorkbook(out)) {
            final XSSFSheet sheet = workbook.getSheetAt(0);
            row = sheet.getRow(0);
            assertCellValue(0, ColumnCode.A.text);
            assertCellValue(1, ColumnCode.B.text);
        }
    }

    @Test
    void annotatedNameMustHaveAnnotation() throws Exception {
        ColumnNameResolver byText = AnnotatedNameResolver.forValue(NameOf.class, ColumnCode::getText);
        final SpreadsheetMapper mapper = new SpreadsheetMapper()
                .setColumnNameResolver(byText);
        assertThatThrownBy(() -> mapper.writeValue(out, null, MissingAnnotationEntity.class))
                .isInstanceOf(InvalidDefinitionException.class)
                .hasMessageContaining("Annotation `@%s` must not be null for property", NameOf.class.getSimpleName());
    }

    void assertCellValue(final int cellnum, String expected) {
        assertThat(row.getCell(cellnum).getStringCellValue()).isEqualTo(expected);
    }

    void assertCellValue(final int cellnum, double expected) {
        assertThat(row.getCell(cellnum).getNumericCellValue()).isEqualTo(expected);
    }

    @Getter
    @RequiredArgsConstructor
    public enum ColumnCode {
        A("Code A"), B("Code B");
        final String text;
        // ...
    }

    @Retention(RetentionPolicy.RUNTIME)
    @interface NameOf {
        ColumnCode value();
    }

    @Data
    @DataGrid
    class AnnotatedEntity {
        @NameOf(ColumnCode.A)
        @DataColumn("It will be overwritten")
        int a;
        @NameOf(ColumnCode.B)
        int b;
    }

    @Data
    @DataGrid
    class MissingAnnotationEntity {
        @DataColumn("Annotation `@NameOf` must not be null")
        int a;
        @NameOf(ColumnCode.B)
        int b;
    }
}
