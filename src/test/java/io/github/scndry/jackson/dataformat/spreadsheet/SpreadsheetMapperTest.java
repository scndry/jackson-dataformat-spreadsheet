package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;

import static org.assertj.core.api.Assertions.*;

class SpreadsheetMapperTest {

    @TempDir File tempDir;
    SpreadsheetMapper mapper;
    File fixture;

    @DataGrid
    static class Entry {
        public int a;
        public int b;
        public Entry() {}
        public Entry(int a, int b) { this.a = a; this.b = b; }

        @Override
        public boolean equals(Object o) {
            if (!(o instanceof Entry)) return false;
            Entry e = (Entry) o;
            return a == e.a && b == e.b;
        }

        @Override
        public int hashCode() {
            return 31 * a + b;
        }
    }

    @BeforeEach
    void setUp() throws Exception {
        mapper = new SpreadsheetMapper();
        fixture = createFixture();
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

    // -- Read --

    @Test
    void readValue() throws Exception {
        Entry value = mapper.readValue(fixture, Entry.class);
        assertThat(value).isEqualTo(new Entry(1, 2));
    }

    @Test
    void readValues() throws Exception {
        List<Entry> values = mapper.readValues(fixture, Entry.class);
        assertThat(values).containsExactly(
                new Entry(1, 2), new Entry(3, 4));
    }

    @Test
    void readValuesBySheetName() throws Exception {
        SheetInput<File> src = SheetInput.source(fixture, "Entries");
        List<Entry> values = mapper.readValues(src, Entry.class);
        assertThat(values).containsExactly(
                new Entry(1, 2), new Entry(3, 4));
    }

    @Test
    void readValuesFromSheetObject() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(fixture)) {
            List<Entry> values = mapper.readValues(
                    wb.getSheetAt(0), Entry.class);
            assertThat(values).containsExactly(
                    new Entry(1, 2), new Entry(3, 4));
        }
    }

    // -- Read (error) --

    @Test
    void readValuesNoSheetThrows() {
        SheetInput<File> src = SheetInput.source(fixture, "NoSheet");
        assertThatThrownBy(() -> mapper.readValues(src, Entry.class))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageStartingWith("No sheet for");
    }

    @Test
    void readValuesOutOfRangeThrows() {
        SheetInput<File> src = SheetInput.source(fixture, 1);
        assertThatThrownBy(() -> mapper.readValues(src, Entry.class))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("out of range");
    }

    // -- Write --

    @Test
    void writeAndReadBack() throws Exception {
        File out = new File(tempDir, "out.xlsx");
        mapper.writeValue(out, new Entry(1, 2));
        Entry actual = mapper.readValue(out, Entry.class);
        assertThat(actual).isEqualTo(new Entry(1, 2));
    }

    @Test
    void writeCollectionAndReadBack() throws Exception {
        File out = new File(tempDir, "out.xlsx");
        List<Entry> data = java.util.Arrays.asList(
                new Entry(1, 2), new Entry(3, 4));
        mapper.writeValue(out, data, Entry.class);
        List<Entry> actual = mapper.readValues(out, Entry.class);
        assertThat(actual).containsExactly(
                new Entry(1, 2), new Entry(3, 4));
    }

    @Test
    void writeWithSheetName() throws Exception {
        File out = new File(tempDir, "out.xlsx");
        List<Entry> data = java.util.Arrays.asList(
                new Entry(1, 2), new Entry(3, 4));
        SheetOutput<File> output = SheetOutput.target(out, "Entries");
        mapper.writeValue(output, data, Entry.class);
        try (XSSFWorkbook wb = new XSSFWorkbook(out)) {
            assertThat(wb.getSheet("Entries")).isNotNull();
        }
    }

    // -- Write (error) --

    @Test
    void writeCollectionWithoutTypeThrows() {
        File out = new File(tempDir, "out.xlsx");
        List<Entry> data = java.util.Arrays.asList(
                new Entry(1, 2));
        assertThatThrownBy(() -> mapper.writeValue(out, data))
                .isInstanceOf(IllegalArgumentException.class);
    }
}
