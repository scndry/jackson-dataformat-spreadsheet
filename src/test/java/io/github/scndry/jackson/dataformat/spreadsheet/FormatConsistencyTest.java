package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Layer 3: SSML (streaming) vs POI (in-memory) consistency.
 * Write a file, then read via both paths and verify identical results.
 */
class FormatConsistencyTest {

    SpreadsheetMapper mapper;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Record {
        String text;
        int num;
        double decimal;
        boolean flag;
    }

    @Test
    void ssmlAndPoiProduceSameResult() throws Exception {
        File file = tempFile("consistency.xlsx");
        List<Record> input = Arrays.asList(
                new Record("hello", 1, 1.5, true),
                new Record("world", 2, 2.5, false),
                new Record("test", -1, 0.0, true)
        );

        mapper.writeValue(file, input, Record.class);

        // SSML path (streaming) — reads from File via SSMLSheetReader
        List<Record> ssmlResult = mapper.readValues(file, Record.class);

        // POI path (in-memory) — reads from POI Sheet via POISheetReader
        List<Record> poiResult;
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            poiResult = mapper.readValues(sheet, Record.class);
        }

        assertThat(ssmlResult).isEqualTo(input);
        assertThat(poiResult).isEqualTo(input);
        assertThat(ssmlResult).isEqualTo(poiResult);
    }

    @Test
    void ssmlViaInputStream() throws Exception {
        File file = tempFile("stream.xlsx");
        List<Record> input = Arrays.asList(
                new Record("stream", 42, 3.14, false)
        );

        mapper.writeValue(file, input, Record.class);

        // File path
        List<Record> fileResult = mapper.readValues(file, Record.class);

        // InputStream path
        List<Record> streamResult;
        try (InputStream is = new FileInputStream(file)) {
            streamResult = mapper.readValues(is, Record.class);
        }

        assertThat(fileResult).isEqualTo(streamResult);
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class WithNull {
        String present;
        String absent;
    }

    @Test
    void nullFieldConsistency() throws Exception {
        File file = tempFile("nulls.xlsx");

        List<WithNull> input = Arrays.asList(
                new WithNull("yes", null),
                new WithNull("also", "here")
        );

        mapper.writeValue(file, input, WithNull.class);

        List<WithNull> ssmlResult = mapper.readValues(file, WithNull.class);

        List<WithNull> poiResult;
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            poiResult = mapper.readValues(wb.getSheetAt(0), WithNull.class);
        }

        assertThat(ssmlResult).isEqualTo(poiResult);
    }

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }
}
