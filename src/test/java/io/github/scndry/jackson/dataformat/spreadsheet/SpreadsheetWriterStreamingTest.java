package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.databind.SequenceWriter;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Streaming write through {@link SpreadsheetWriter#writeValues(Sheet)} and
 * {@link SpreadsheetWriter#writeValues(SheetOutput)} — the {@link SequenceWriter}
 * pattern used when rows arrive lazily (paged repository, JDBC cursor) and the
 * caller must release each row's memory before the next is produced.
 *
 * <p>The verification chain is write-many-incremental → read-back-all to assert
 * row ordering and count survive the streaming path end-to-end.</p>
 */
class SpreadsheetWriterStreamingTest {

    @TempDir File tempDir;
    SpreadsheetMapper mapper;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Item {
        @DataColumn("ID") private long id;
        @DataColumn("Name") private String name;
    }

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    @Test
    void writeValuesToSheet_streamsItemsIncrementally() throws Exception {
        File file = new File(tempDir, "stream-sheet.xlsx");
        SpreadsheetWriter writer = mapper.sheetWriterFor(Item.class);

        try (SXSSFWorkbook wb = new SXSSFWorkbook();
             OutputStream os = new FileOutputStream(file)) {
            Sheet sheet = wb.createSheet("Stream");
            try (SequenceWriter seq = writer.writeValues(sheet)) {
                for (int i = 0; i < 50; i++) {
                    seq.write(new Item(i, "row-" + i));
                }
            }
            wb.write(os);
            wb.dispose();
        }

        List<Item> read = mapper.readValues(file, Item.class);
        assertThat(read).hasSize(50);
        assertThat(read.get(0).getId()).isZero();
        assertThat(read.get(49).getName()).isEqualTo("row-49");
    }

    @Test
    void writeValuesToSheetOutput_streamsItemsIncrementally() throws Exception {
        File file = new File(tempDir, "stream-sheetoutput.xlsx");
        SpreadsheetWriter writer = mapper.sheetWriterFor(Item.class);

        try (SequenceWriter seq = writer.writeValues(SheetOutput.target(file, "Stream"))) {
            for (int i = 0; i < 50; i++) {
                seq.write(new Item(i, "row-" + i));
            }
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheet("Stream")).isNotNull();
        }
        List<Item> read = mapper.readValues(file, Item.class);
        assertThat(read).hasSize(50);
        assertThat(read.get(25).getId()).isEqualTo(25);
        assertThat(read.get(25).getName()).isEqualTo("row-25");
    }
}
