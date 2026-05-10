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
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Verifies how BackWriteProjection.hasOuterFieldAfterList behaves when the
 * column order is altered by @JsonPropertyOrder.
 *
 * The detection scans SpreadsheetSchema._columns (spreadsheet column order).
 * This test asks: does the spreadsheet column order follow declaration order,
 * or does it follow @JsonPropertyOrder? If the latter, detection is precise
 * for the @JsonPropertyOrder case; if the former, detection may produce
 * false positives when @JsonPropertyOrder moves the outer field before the
 * list.
 */
class BackWriteDetectionTest {

    private static final Path DEBUG_OUTPUT_DIR = Paths.get("build/debug-output");
    private static final String LIMIT_PROP = "spreadsheet.backWriteBufferBytes";
    private String _savedLimit;

    @BeforeEach
    void saveLimit() {
        _savedLimit = System.getProperty(LIMIT_PROP);
    }

    @AfterEach
    void restoreLimit() {
        if (_savedLimit == null) {
            System.clearProperty(LIMIT_PROP);
        } else {
            System.setProperty(LIMIT_PROP, _savedLimit);
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Item {
        @DataColumn("v") int v;
    }

    // Declaration order: id, items, total — outer-after-list pattern
    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class BasicRecord {
        @DataColumn("id") int id;
        List<Item> items;
        @DataColumn("total") int total;
    }

    // Declaration order: id, items, total
    // @JsonPropertyOrder forces: id, total, items — outer-first
    // If Jackson follows @JsonPropertyOrder, back-write does NOT occur and
    // detection ideally should return false (no fail-fast).
    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    @JsonPropertyOrder({"id", "total", "items"})
    static class ReorderedRecord {
        @DataColumn("id") int id;
        List<Item> items;
        @DataColumn("total") int total;
    }

    @Test
    void declarationOrder_outerAfterList_triggersFailFast() throws Exception {
        // Baseline: declaration order = outer-after-list. With a tiny limit,
        // the projected check must fire.
        System.setProperty(LIMIT_PROP, String.valueOf(16 * 1024));

        List<Item> items = new ArrayList<>(1000);
        for (int i = 0; i < 1000; i++) items.add(new Item(i));
        BasicRecord record = new BasicRecord(1, items, 99);

        File file = _debugFile("detect-basic.xlsx");
        assertThatThrownBy(() ->
                new SpreadsheetMapper().writeValue(file,
                        Collections.singletonList(record), BasicRecord.class))
                .hasMessageContaining("projected back-write buffer");
    }

    @Test
    void jsonPropertyOrder_outerFirst_writesNormally() throws Exception {
        // @JsonPropertyOrder reorders to outer-first. If schema column order
        // reflects @JsonPropertyOrder, detection should return false and the
        // write should complete normally — the tiny limit must not trip.
        System.setProperty(LIMIT_PROP, String.valueOf(16 * 1024));

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
