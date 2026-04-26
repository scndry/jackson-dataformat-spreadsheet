package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.*;

class WriterReaderApiTest {

    @TempDir File tempDir;
    SpreadsheetMapper mapper;

    @DataGrid
    static class Item {
        public String name;
        public int qty;
        public Item() {}
        public Item(String name, int qty) { this.name = name; this.qty = qty; }
    }

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    @Test
    void writeToOutputStream() throws Exception {
        File file = new File(tempDir, "out.xlsx");
        List<Item> items = Arrays.asList(new Item("A", 1), new Item("B", 2));
        try (OutputStream os = new FileOutputStream(file)) {
            mapper.writeValue(os, items, Item.class);
        }
        List<Item> read = mapper.readValues(file, Item.class);
        assertThat(read).hasSize(2);
        assertThat(read.get(0).name).isEqualTo("A");
    }

    @Test
    void writeToSheetOutput() throws Exception {
        File file = new File(tempDir, "sheetout.xlsx");
        SheetOutput<File> output = SheetOutput.target(file, "MySheet");
        mapper.writeValue(output, Arrays.asList(new Item("X", 10)), Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheet("MySheet")).isNotNull();
            assertThat(wb.getSheet("MySheet").getRow(1).getCell(0).getStringCellValue()).isEqualTo("X");
        }
    }

    @Test
    void writeToSheetDirectly() throws Exception {
        File file = new File(tempDir, "direct.xlsx");
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Direct");
            mapper.writeValue(sheet, Arrays.asList(new Item("D", 5)), Item.class);
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
            wb.dispose();
        }
        List<Item> read = mapper.readValues(SheetInput.source(file, "Direct"), Item.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0).name).isEqualTo("D");
    }

    @Test
    void readFromInputStream() throws Exception {
        File file = new File(tempDir, "stream.xlsx");
        mapper.writeValue(file, Arrays.asList(new Item("S", 3)), Item.class);

        try (InputStream is = new FileInputStream(file)) {
            List<Item> read = mapper.readValues(is, Item.class);
            assertThat(read).hasSize(1);
            assertThat(read.get(0).name).isEqualTo("S");
        }
    }

    @Test
    void readFromSheetInput() throws Exception {
        File file = new File(tempDir, "input.xlsx");
        SheetOutput<File> output = SheetOutput.target(file, "Data");
        mapper.writeValue(output, Arrays.asList(new Item("I", 7)), Item.class);

        SheetInput<File> input = SheetInput.source(file, "Data");
        List<Item> read = mapper.readValues(input, Item.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0).name).isEqualTo("I");
    }

    @Test
    void readFromSheetDirectly() throws Exception {
        File file = new File(tempDir, "poi.xlsx");
        mapper.writeValue(file, Arrays.asList(new Item("P", 9)), Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Item item = mapper.readValue(sheet, Item.class);
            assertThat(item.name).isEqualTo("P");
            assertThat(item.qty).isEqualTo(9);
        }
    }

    @Test
    void sheetWriterFor() throws Exception {
        File file = new File(tempDir, "writer.xlsx");
        SpreadsheetWriter writer = mapper.sheetWriterFor(Item.class);
        writer.writeValues(file)
                .write(new Item("W1", 1))
                .write(new Item("W2", 2))
                .close();

        List<Item> read = mapper.readValues(file, Item.class);
        assertThat(read).hasSize(2);
    }

    @Test
    void sheetReaderFor() throws Exception {
        File file = new File(tempDir, "reader.xlsx");
        mapper.writeValue(file, Arrays.asList(new Item("R", 4)), Item.class);

        SpreadsheetReader reader = mapper.sheetReaderFor(Item.class);
        try (SheetMappingIterator<Item> iter = reader.readValues(file)) {
            assertThat(iter.hasNext()).isTrue();
            Item item = iter.next();
            assertThat(item.name).isEqualTo("R");
        }
    }

    @Test
    void readExceptionLocationIsSheetLocation() throws Exception {
        File file = new File(tempDir, "bad-type.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0).createCell(0).setCellValue("name");
            sheet.createRow(0).createCell(1).setCellValue("qty");
            org.apache.poi.ss.usermodel.Row row = sheet.createRow(1);
            row.createCell(0).setCellValue("OK");
            row.createCell(1).setCellValue("not-a-number");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        assertThatThrownBy(() -> mapper.readValues(file, Item.class))
                .satisfies(e -> {
                    SheetLocation loc = SheetLocation.of((Exception) e);
                    assertThat(loc).isNotNull();
                    assertThat(loc.getRow()).isGreaterThanOrEqualTo(1);
                });
    }

    @Test
    void writeValueAsBytes() throws Exception {
        byte[] bytes = mapper.writeValueAsBytes(new Item("B", 8), Item.class);
        assertThat(bytes).isNotEmpty();

        try (InputStream is = new ByteArrayInputStream(bytes)) {
            Item item = mapper.readValue(is, Item.class);
            assertThat(item.name).isEqualTo("B");
        }
    }
}
