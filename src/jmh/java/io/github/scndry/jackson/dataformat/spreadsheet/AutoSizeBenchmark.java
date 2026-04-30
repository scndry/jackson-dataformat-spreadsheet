package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.annotation.OptBoolean;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Measures inline {@code SheetUtil.getCellWidth} overhead in {@link POISheetWriter#_write}
 * when {@code autoSize} is enabled. Compares against the baseline path with {@code autoSize}
 * disabled to isolate measurement cost.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class AutoSizeBenchmark {

    @Param({"1000", "10000", "100000"})
    int rowCount;

    List<NoAutoSize> baseline;
    List<WithAutoSize> autoSized;
    File file;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class NoAutoSize {
        public String name;
        public int quantity;
        public double price;
        public String description;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(autoSizeColumn = OptBoolean.TRUE)
    public static class WithAutoSize {
        public String name;
        public int quantity;
        public double price;
        public String description;
    }

    @Setup(Level.Trial)
    public void setUp() {
        baseline = new ArrayList<>(rowCount);
        autoSized = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            baseline.add(new NoAutoSize("item-" + i, i, i * 1.5, "description of item " + i));
            autoSized.add(new WithAutoSize("item-" + i, i, i * 1.5, "description of item " + i));
        }
    }

    @Setup(Level.Invocation)
    public void setUpFile() throws IOException {
        file = File.createTempFile("bench-autosize-", ".xlsx");
        file.deleteOnExit();
    }

    @TearDown(Level.Invocation)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void poiWithoutAutoSize(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapper.writeValue(file, baseline, NoAutoSize.class);
        bh.consume(file);
    }

    @Benchmark
    public void poiWithAutoSize(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapper.writeValue(file, autoSized, WithAutoSize.class);
        bh.consume(file);
    }
}
