package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

/**
 * Cell-count-parity benchmark — sizes flat row count and nested inner
 * row count so both fixtures hold roughly the same number of written
 * cells. Removes the "file-size parity" framing (same row count, half
 * the cells) and surfaces the V2 record-tree algorithm's per-cell cost
 * relative to the flat path.
 *
 * <p>Flat fixture: targetCells / 10 rows (10 cols per row).
 * Nested fixture: round_to_10(targetCells / 4.6) inner rows
 * (itemsPerOuter = 10 → 4 inner + 6 outer × (1/10) cells per inner row).
 * Cell counts match within 0.005 %.
 *
 * <p>Other libraries are omitted (no nested-list support).
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class NestedCellParityBenchmark {

    @Param({"100000", "500000", "1000000"})
    int targetCells;

    File nestedFile;
    File flatFile;
    List<NestedRow> nestedData;
    List<BenchRow> flatData;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoiRead;
    SpreadsheetMapper mapperPoiWrite;

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        final int flatRows = targetCells / 10;
        final int nestedInner = Math.round(targetCells / 4.6f / 10) * 10;

        mapper = new SpreadsheetMapper();
        mapperPoiRead = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapperPoiWrite = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));

        flatFile = BenchWorkbookBuilder.createSampleFile("bench-parity-flat-", flatRows);
        nestedFile = NestedBenchWorkbookBuilder.createSampleFile("bench-parity-nested-", nestedInner);

        final int outerCount = nestedInner / NestedBenchWorkbookBuilder.ITEMS_PER_OUTER;
        nestedData = NestedRow.sample(outerCount, NestedBenchWorkbookBuilder.ITEMS_PER_OUTER);
        flatData = new ArrayList<>(flatRows);
        for (int i = 0; i < flatRows; i++) {
            flatData.add(BenchRow.create(i));
        }
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        nestedFile.delete();
        flatFile.delete();
    }

    @Benchmark
    public void flatRead(Blackhole bh) throws IOException {
        List<BenchRow> values = mapper.readValues(flatFile, BenchRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flatReadPoi(Blackhole bh) throws IOException {
        List<BenchRow> values = mapperPoiRead.readValues(flatFile, BenchRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flatWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, flatData, BenchRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void flatWritePoi(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapperPoiWrite.writeValue(out, flatData, BenchRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void nestedRead(Blackhole bh) throws IOException {
        List<NestedRow> values = mapper.readValues(nestedFile, NestedRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void nestedReadPoi(Blackhole bh) throws IOException {
        List<NestedRow> values = mapperPoiRead.readValues(nestedFile, NestedRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void nestedWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, nestedData, NestedRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void nestedWritePoi(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapperPoiWrite.writeValue(out, nestedData, NestedRow.class);
        bh.consume(out);
    }
}
