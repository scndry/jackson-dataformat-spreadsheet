package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

/**
 * Nested-list write benchmark — outer-after-list shape exercises the
 * SSML back-write path (totalAmount/createdAt sit after items[]).
 * BenchRow flat baseline is included for direct cross-shape comparison
 * within the same JVM/trial (same total inner row count = same file
 * size). Other writers are omitted because they do not emit nested-list
 * layouts.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class NestedWriteBenchmark {

    @Param({"10000", "50000", "100000"})
    int rowCount;

    List<NestedRow> nestedData;
    List<BenchRow> flatData;
    File file;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoi;

    @Setup(Level.Trial)
    public void setUp() {
        mapper = new SpreadsheetMapper();
        mapperPoi = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        final int outerCount = rowCount / NestedBenchWorkbookBuilder.ITEMS_PER_OUTER;
        nestedData = NestedRow.sample(outerCount, NestedBenchWorkbookBuilder.ITEMS_PER_OUTER);
        flatData = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            flatData.add(BenchRow.create(i));
        }
    }

    @Setup(Level.Invocation)
    public void setUpFile() throws IOException {
        file = File.createTempFile("bench-nested-write-", ".xlsx");
        file.deleteOnExit();
    }

    @TearDown(Level.Invocation)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void nested(Blackhole bh) throws IOException {
        mapper.writeValue(file, nestedData, NestedRow.class);
        bh.consume(file);
    }

    @Benchmark
    public void nestedPoi(Blackhole bh) throws IOException {
        mapperPoi.writeValue(file, nestedData, NestedRow.class);
        bh.consume(file);
    }

    @Benchmark
    public void flat(Blackhole bh) throws IOException {
        mapper.writeValue(file, flatData, BenchRow.class);
        bh.consume(file);
    }

    @Benchmark
    public void flatPoi(Blackhole bh) throws IOException {
        mapperPoi.writeValue(file, flatData, BenchRow.class);
        bh.consume(file);
    }
}
