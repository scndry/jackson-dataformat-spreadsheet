package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Covers {@link SpreadsheetReader} overloads not exercised through
 * {@link SpreadsheetMapper}'s top-level methods — the {@code Sheet} /
 * {@link SheetInput} variants of {@code readValue} / {@code readValues} /
 * {@code createParser}, the empty-file null path, and fluent factories
 * ({@link SpreadsheetReader#forType}, {@link SpreadsheetReader#parserFactory},
 * the {@code with(...)} chain).
 */
class SpreadsheetReaderTest {

    @TempDir File tempDir;

    @DataGrid
    static class Item {
        public String name;
        public int qty;
        public Item() {}
        public Item(String name, int qty) { this.name = name; this.qty = qty; }
    }

    SpreadsheetMapper mapper;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    private File writeOne(String filename, String name, int qty) throws Exception {
        File file = new File(tempDir, filename);
        mapper.writeValue(file, Arrays.asList(new Item(name, qty)), Item.class);
        return file;
    }

    @Test
    void readValue_File_singleRow() throws Exception {
        File file = writeOne("rv-file.xlsx", "F", 1);
        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        Item item = reader.readValue(file);
        assertThat(item.name).isEqualTo("F");
    }

    @Test
    void readValue_InputStream_singleRow() throws Exception {
        File file = writeOne("rv-is.xlsx", "I", 2);
        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        try (InputStream is = new FileInputStream(file)) {
            Item item = reader.readValue(is);
            assertThat(item.name).isEqualTo("I");
        }
    }

    @Test
    void readValue_Sheet_singleRow() throws Exception {
        File file = writeOne("rv-sheet.xlsx", "S", 3);
        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Item item = reader.readValue(wb.getSheetAt(0));
            assertThat(item.name).isEqualTo("S");
        }
    }

    @Test
    void readValue_SheetInput_singleRow() throws Exception {
        File file = new File(tempDir, "rv-input.xlsx");
        mapper.writeValue(SheetOutput.target(file, "Data"),
                Arrays.asList(new Item("N", 4)), Item.class);

        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        Item item = reader.readValue(SheetInput.source(file, "Data"));
        assertThat(item.name).isEqualTo("N");
    }

    @Test
    void readValues_Sheet_iteration() throws Exception {
        File file = new File(tempDir, "rvs-sheet.xlsx");
        mapper.writeValue(file, Arrays.asList(new Item("A", 1), new Item("B", 2)), Item.class);

        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        try (XSSFWorkbook wb = new XSSFWorkbook(file);
             SheetMappingIterator<Item> iter = reader.readValues(wb.getSheetAt(0))) {
            assertThat(iter.next().name).isEqualTo("A");
            assertThat(iter.next().name).isEqualTo("B");
        }
    }

    @Test
    void readValues_SheetInput_iteration() throws Exception {
        File file = new File(tempDir, "rvs-input.xlsx");
        mapper.writeValue(SheetOutput.target(file, "Data"),
                Arrays.asList(new Item("X", 1), new Item("Y", 2)), Item.class);

        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        try (SheetMappingIterator<Item> iter = reader.readValues(SheetInput.source(file, "Data"))) {
            assertThat(iter.next().name).isEqualTo("X");
            assertThat(iter.next().name).isEqualTo("Y");
        }
    }

    @Test
    void createParser_Sheet_returnsSheetParser() throws Exception {
        File file = writeOne("cp-sheet.xlsx", "P", 1);
        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            try (SheetParser parser = reader.createParser(wb.getSheetAt(0))) {
                assertThat(parser).isNotNull();
            }
        }
    }

    @Test
    void createParser_SheetInput_returnsSheetParser() throws Exception {
        File file = new File(tempDir, "cp-input.xlsx");
        mapper.writeValue(SheetOutput.target(file, "Data"),
                Arrays.asList(new Item("P", 1)), Item.class);

        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        try (SheetParser parser = reader.createParser(SheetInput.source(file, "Data"))) {
            assertThat(parser).isNotNull();
        }
    }

    @Test
    void emptyFile_readValue_returnsNull() throws Exception {
        File file = new File(tempDir, "empty.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0).createCell(0).setCellValue("name");
            sheet.createRow(0).createCell(1).setCellValue("qty");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        Item item = reader.readValue(file);
        assertThat(item).isNull();
    }

    @Test
    void forType_returnsSpreadsheetReader() throws Exception {
        SpreadsheetReader base = (SpreadsheetReader) mapper.reader();
        SpreadsheetReader typed = base.forType(Item.class);
        assertThat(typed).isInstanceOf(SpreadsheetReader.class);
    }

    @Test
    void parserFactory_returnsSpreadsheetFactory() throws Exception {
        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        SpreadsheetFactory factory = reader.parserFactory();
        assertThat(factory).isInstanceOf(SpreadsheetFactory.class);
    }

    @Test
    void withFeatures_chainCreatesNewReader() throws Exception {
        SpreadsheetReader base = mapper.sheetReaderFor(Item.class);
        SpreadsheetReader configured = (SpreadsheetReader) base.with(base.getConfig());
        assertThat(configured).isInstanceOf(SpreadsheetReader.class);
    }
}
