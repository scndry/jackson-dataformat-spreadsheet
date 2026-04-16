package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.databind.MappingIterator;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Layer 1: Integration round-trip tests.
 * Write → Read → verify original equals result.
 */
class RoundTripTest {

    SpreadsheetMapper mapper;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    // -- Fixture types --

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class BasicTypes {
        String text;
        int intVal;
        double doubleVal;
        boolean boolVal;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class NestedPojo {
        String name;
        Address address;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    static class Address {
        String city;
        int zip;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class NullableFields {
        String required;
        String optional;
    }

    // -- Tests --

    @Test
    void basicTypes() throws Exception {
        File file = tempFile("basic.xlsx");
        List<BasicTypes> input = Arrays.asList(
                new BasicTypes("hello", 42, 3.14, true),
                new BasicTypes("world", -1, 0.0, false)
        );

        mapper.writeValue(file, input, BasicTypes.class);
        List<BasicTypes> output = mapper.readValues(file, BasicTypes.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void singleRow() throws Exception {
        File file = tempFile("single.xlsx");
        BasicTypes input = new BasicTypes("one", 1, 1.0, true);

        mapper.writeValue(file, input);
        BasicTypes output = mapper.readValue(file, BasicTypes.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void nestedPojo() throws Exception {
        File file = tempFile("nested.xlsx");
        List<NestedPojo> input = Arrays.asList(
                new NestedPojo("Alice", new Address("Seoul", 12345)),
                new NestedPojo("Bob", new Address("Busan", 67890))
        );

        mapper.writeValue(file, input, NestedPojo.class);
        List<NestedPojo> output = mapper.readValues(file, NestedPojo.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void nullField() throws Exception {
        File file = tempFile("nullable.xlsx");
        List<NullableFields> input = Arrays.asList(
                new NullableFields("present", "also present"),
                new NullableFields("present", null)
        );

        mapper.writeValue(file, input, NullableFields.class);
        List<NullableFields> output = mapper.readValues(file, NullableFields.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void emptyList() throws Exception {
        File file = tempFile("empty.xlsx");
        List<BasicTypes> input = Arrays.asList();

        mapper.writeValue(file, input, BasicTypes.class);
        List<BasicTypes> output = mapper.readValues(file, BasicTypes.class);

        assertThat(output).isEmpty();
    }

    @Test
    void sheetName() throws Exception {
        File file = tempFile("named.xlsx");
        List<BasicTypes> input = Arrays.asList(
                new BasicTypes("test", 1, 2.0, true)
        );

        mapper.writeValue(SheetOutput.target(file, "MySheet"), input, BasicTypes.class);
        List<BasicTypes> output = mapper.readValues(
                SheetInput.source(file, "MySheet"), BasicTypes.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void sheetNameNotFound() throws Exception {
        File file = tempFile("named2.xlsx");
        mapper.writeValue(SheetOutput.target(file, "Exists"), Arrays.asList(
                new BasicTypes("test", 1, 2.0, true)), BasicTypes.class);

        assertThatThrownBy(() ->
                mapper.readValues(SheetInput.source(file, "DoesNotExist"), BasicTypes.class))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void poiSheetDirectAccess() throws Exception {
        List<BasicTypes> input = Arrays.asList(
                new BasicTypes("direct", 99, 1.5, false)
        );

        // Write via POI Sheet
        try (Workbook wb = new SXSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Direct");
            mapper.writeValue(sheet, input, BasicTypes.class);
            File file = tempFile("direct.xlsx");
            try (OutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }

            // Read via POI Sheet
            try (XSSFWorkbook readWb = new XSSFWorkbook(file)) {
                Sheet readSheet = readWb.getSheet("Direct");
                List<BasicTypes> output = mapper.readValues(readSheet, BasicTypes.class);
                assertThat(output).isEqualTo(input);
            }
        }
    }

    @Test
    void originOffset() throws Exception {
        File file = tempFile("origin.xlsx");
        SpreadsheetMapper offsetMapper = SpreadsheetMapper.builder()
                .origin("B3")
                .build();

        List<BasicTypes> input = Arrays.asList(
                new BasicTypes("offset", 10, 2.5, true)
        );

        offsetMapper.writeValue(file, input, BasicTypes.class);
        List<BasicTypes> output = offsetMapper.readValues(file, BasicTypes.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void mappingIteratorStreaming() throws Exception {
        File file = tempFile("streaming.xlsx");
        List<BasicTypes> input = Arrays.asList(
                new BasicTypes("a", 1, 1.0, true),
                new BasicTypes("b", 2, 2.0, false),
                new BasicTypes("c", 3, 3.0, true)
        );

        mapper.writeValue(file, input, BasicTypes.class);

        // Read one by one via MappingIterator
        try (MappingIterator<BasicTypes> iterator =
                     mapper.sheetReaderFor(BasicTypes.class).readValues(file)) {
            assertThat(iterator.next()).isEqualTo(input.get(0));
            assertThat(iterator.next()).isEqualTo(input.get(1));
            assertThat(iterator.next()).isEqualTo(input.get(2));
            assertThat(iterator.hasNext()).isFalse();
        }
    }

    @Test
    void collectionWriteRequiresValueType() throws Exception {
        File file = tempFile("notype.xlsx");
        List<BasicTypes> input = Arrays.asList(new BasicTypes("test", 1, 1.0, true));

        assertThatThrownBy(() -> mapper.writeValue(file, input))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("valueType");
    }

    // -- Deep nesting --

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class DeepNested {
        String top;
        Level1 level1;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    static class Level1 {
        String mid;
        Level2 level2;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    static class Level2 {
        int bottom;
    }

    @Test
    void deepNestedPojo() throws Exception {
        File file = tempFile("deep.xlsx");
        List<DeepNested> input = Arrays.asList(
                new DeepNested("a", new Level1("b", new Level2(99)))
        );

        mapper.writeValue(file, input, DeepNested.class);
        List<DeepNested> output = mapper.readValues(file, DeepNested.class);

        assertThat(output).isEqualTo(input);
    }

    // -- useHeader=false + origin --

    @Test
    void originWithoutHeader() throws Exception {
        File file = tempFile("origin-no-header.xlsx");
        SpreadsheetMapper m = SpreadsheetMapper.builder()
                .origin("B2")
                .useHeader(false)
                .build();

        List<BasicTypes> input = Arrays.asList(
                new BasicTypes("offset", 10, 2.5, true)
        );

        m.writeValue(file, input, BasicTypes.class);
        List<BasicTypes> output = m.readValues(file, BasicTypes.class);

        assertThat(output).isEqualTo(input);
    }

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }
}
