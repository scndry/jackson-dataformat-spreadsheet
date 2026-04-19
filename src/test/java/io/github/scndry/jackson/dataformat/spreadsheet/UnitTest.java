package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class UnitTest {

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Simple { String name; int value; }


    @Nested
    class SchemaTest {

        SpreadsheetMapper mapper;

        @BeforeEach
        void setUp() {
            mapper = new SpreadsheetMapper();
        }

        @Test
        void schemaGeneration() throws Exception {
            SpreadsheetSchema schema = mapper.sheetSchemaFor(Simple.class);
            assertThat(schema.getSchemaType()).isEqualTo("spreadsheet");
        }

        @Test
        void dataRowWithHeader() throws Exception {
            SpreadsheetSchema schema = mapper.sheetSchemaFor(Simple.class);
            assertThat(schema.usesHeader()).isTrue();
            assertThat(schema.getDataRow()).isEqualTo(1);
        }

        @Test
        void dataRowWithoutHeader() throws Exception {
            SpreadsheetMapper noHeader = mapper.rebuild().useHeader(false).build();
            SpreadsheetSchema schema = noHeader.sheetSchemaFor(Simple.class);
            assertThat(schema.usesHeader()).isFalse();
            assertThat(schema.getDataRow()).isEqualTo(0);
        }

        @Test
        void nonDataGridClassRejected() {
            assertThatThrownBy(() -> mapper.sheetSchemaFor(String.class))
                    .hasMessageContaining("@DataGrid");
        }

        @Test
        void collectionTypeRejected() {
            assertThatThrownBy(() -> mapper.sheetSchemaFor(List.class))
                    .hasMessageContaining("Collection or array");
        }
    }

    @Nested
    class UseHeaderTest {

        @TempDir Path tempDir;
        SpreadsheetMapper withHeader;
        SpreadsheetMapper withoutHeader;

        @BeforeEach
        void setUp() {
            withHeader = new SpreadsheetMapper();
            withoutHeader = SpreadsheetMapper.builder().useHeader(false).build();
        }

        @Test
        void writeAndReadWithHeader() throws Exception {
            File file = tempDir.resolve("with-header.xlsx").toFile();
            List<Simple> input = Arrays.asList(new Simple("test", 42));

            withHeader.writeValue(file, input, Simple.class);
            List<Simple> output = withHeader.readValues(file, Simple.class);

            assertThat(output).isEqualTo(input);
        }

        @Test
        void writeAndReadWithoutHeader() throws Exception {
            File file = tempDir.resolve("no-header.xlsx").toFile();
            List<Simple> input = Arrays.asList(new Simple("test", 42));

            withoutHeader.writeValue(file, input, Simple.class);
            List<Simple> output = withoutHeader.readValues(file, Simple.class);

            assertThat(output).isEqualTo(input);
        }

        @Test
        void withoutHeaderMultipleRows() throws Exception {
            File file = tempDir.resolve("no-header-multi.xlsx").toFile();
            List<Simple> input = Arrays.asList(
                    new Simple("first", 1),
                    new Simple("second", 2)
            );

            withoutHeader.writeValue(file, input, Simple.class);
            List<Simple> output = withoutHeader.readValues(file, Simple.class);

            assertThat(output).isEqualTo(input);
        }
    }

}
