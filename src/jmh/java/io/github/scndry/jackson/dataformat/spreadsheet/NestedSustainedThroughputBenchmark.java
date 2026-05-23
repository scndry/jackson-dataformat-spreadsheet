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
 * Sustained-throughput benchmark for the nested-list read/write path.
 * Mirrors {@link SustainedThroughputBenchmark}'s 60-second measurement
 * window to surface steady-state cost (allocation rate, GC pause
 * accumulation) that single-shot per-op figures hide. Includes the
 * flat BenchRow baselines for direct cross-shape comparison within
 * the same trial. Other libraries are omitted because they do not
 * read/write nested lists.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.Throughput)
@OutputTimeUnit(TimeUnit.SECONDS)
@Warmup(iterations = 1, time = 10)
@Measurement(iterations = 1, time = 60)
@Fork(1)
public class NestedSustainedThroughputBenchmark {

    static final int ROW_COUNT = 10_000;

    File nestedFile;
    File flatFile;
    List<NestedRow> nestedData;
    List<BenchRow> flatData;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoiRead;
    SpreadsheetMapper mapperPoiWrite;

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        mapper = new SpreadsheetMapper();
        mapperPoiRead = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapperPoiWrite = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        nestedFile = NestedBenchWorkbookBuilder.createSampleFile("bench-sustained-nested-", ROW_COUNT);
        flatFile = BenchWorkbookBuilder.createSampleFile("bench-sustained-flat-", ROW_COUNT);

        final int outerCount = ROW_COUNT / NestedBenchWorkbookBuilder.ITEMS_PER_OUTER;
        nestedData = NestedRow.sample(outerCount, NestedBenchWorkbookBuilder.ITEMS_PER_OUTER);
        flatData = new ArrayList<>(ROW_COUNT);
        for (int i = 0; i < ROW_COUNT; i++) {
            flatData.add(BenchRow.create(i));
        }
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        nestedFile.delete();
        flatFile.delete();
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
}
