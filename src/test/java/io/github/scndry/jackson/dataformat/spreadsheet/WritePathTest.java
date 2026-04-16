package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Write path tests — only API entries not covered by RoundTripTest.
 */
class WritePathTest {

    SpreadsheetMapper mapper;
    @TempDir Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Row { String name; int value; }

    @Test
    void writeToOutputStream() throws Exception {
        File file = tempDir.resolve("stream.xlsx").toFile();
        try (OutputStream os = new FileOutputStream(file)) {
            mapper.writeValue(os, new Row("stream", 2));
        }
        Row result = mapper.readValue(file, Row.class);
        assertThat(result.getName()).isEqualTo("stream");
    }

    @Test
    void writeAsBytes() throws Exception {
        byte[] bytes = mapper.writeValueAsBytes(new Row("bytes", 1), Row.class);
        assertThat(bytes).isNotEmpty();
        Row result = mapper.readValue(new ByteArrayInputStream(bytes), Row.class);
        assertThat(result.getName()).isEqualTo("bytes");
    }

    @Test
    void writeArrayWithoutValueTypeThrows() {
        File file = tempDir.resolve("arr.xlsx").toFile();
        assertThatThrownBy(() -> mapper.writeValue(file, new Row[]{new Row("x", 1)}))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void writerWithSchema() throws Exception {
        File file = tempDir.resolve("schema.xlsx").toFile();
        SpreadsheetWriter writer = mapper.sheetWriterFor(Row.class);
        writer.writeValue(file, Arrays.asList(new Row("schema", 5)));
        List<Row> result = mapper.readValues(file, Row.class);
        assertThat(result).hasSize(1);
        assertThat(result.get(0).getName()).isEqualTo("schema");
    }
}
