package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonPropertyOrder;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Detection accuracy of the back-write trigger.
 *
 * <p>The trigger inspects {@code SpreadsheetSchema._columns} (spreadsheet
 * column order). When {@code @JsonPropertyOrder} reorders an
 * outer-after-list declaration so the outer field sits <em>before</em> the
 * {@code List<T>}, the gating must return {@code false} — array-scope
 * flush suspension does not engage, the runtime {@code _sb} monitor does
 * not see this scope, and a modest list writes normally.
 */
class BackWriteDetectionTest {

    private static final Path DEBUG_OUTPUT_DIR = Paths.get("build/debug-output");

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Item {
        @DataColumn("v") int v;
    }

    // Declaration order: id, items, total
    // @JsonPropertyOrder forces: id, total, items — outer-first
    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    @JsonPropertyOrder({"id", "total", "items"})
    static class ReorderedRecord {
        @DataColumn("id") int id;
        List<Item> items;
        @DataColumn("total") int total;
    }

    @Test
    void jsonPropertyOrder_outerFirst_writesNormally() throws Exception {
        List<Item> items = new ArrayList<>(1000);
        for (int i = 0; i < 1000; i++) items.add(new Item(i));
        ReorderedRecord record = new ReorderedRecord(1, items, 99);

        File file = _debugFile("detect-reordered.xlsx");
        new SpreadsheetMapper().writeValue(file,
                Collections.singletonList(record), ReorderedRecord.class);

        assertThat(file).exists().isNotEmpty();
    }

    private static File _debugFile(final String name) throws java.io.IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }
}
