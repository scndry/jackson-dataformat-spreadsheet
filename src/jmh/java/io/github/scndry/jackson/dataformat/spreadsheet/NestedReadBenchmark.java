package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

/**
 * Nested-list read benchmark — exercises the V2 record-tree algorithm
 * against a 10-column outer-after-list shape (6 outer + 4 inner,
 * itemsPerOuter=10). BenchRow flat baseline is included for direct
 * cross-shape comparison within the same JVM/trial (same total inner
 * row count = same file size). Other readers are omitted because they
 * do not deserialize nested lists.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class NestedReadBenchmark {

    @Param({"10000", "50000", "100000"})
    int rowCount;

    File nestedFile;
    File flatFile;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoi;

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        mapper = new SpreadsheetMapper();
        mapperPoi = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        nestedFile = NestedBenchWorkbookBuilder.createSampleFile("bench-nested-read-", rowCount);
        flatFile = BenchWorkbookBuilder.createSampleFile("bench-flat-read-", rowCount);
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        nestedFile.delete();
        flatFile.delete();
    }

    @Benchmark
    public void nested(Blackhole bh) throws IOException {
        List<NestedRow> values = mapper.readValues(nestedFile, NestedRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void nestedPoi(Blackhole bh) throws IOException {
        List<NestedRow> values = mapperPoi.readValues(nestedFile, NestedRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flat(Blackhole bh) throws IOException {
        List<BenchRow> values = mapper.readValues(flatFile, BenchRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flatPoi(Blackhole bh) throws IOException {
        List<BenchRow> values = mapperPoi.readValues(flatFile, BenchRow.class);
        bh.consume(values);
    }
}
