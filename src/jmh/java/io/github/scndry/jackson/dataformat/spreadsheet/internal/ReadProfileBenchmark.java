package io.github.scndry.jackson.dataformat.spreadsheet.internal;

import io.github.scndry.jackson.dataformat.spreadsheet.BenchRow;
import io.github.scndry.jackson.dataformat.spreadsheet.BenchWorkbookBuilder;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Internal profiling benchmark — isolates costs: schema generation vs full read
 * (with reused mapper/schema). Used for optimization work; not documented in
 * BENCHMARK.md.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class ReadProfileBenchmark {

    @Param({"1000", "10000", "50000"})
    int rowCount;

    File file;
    SpreadsheetMapper mapper;
    SpreadsheetSchema schema;

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        file = BenchWorkbookBuilder.createSampleFile("bench-profile-", rowCount);
        mapper = new SpreadsheetMapper();
        schema = mapper.sheetSchemaFor(BenchRow.class);
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        file.delete();
    }

    /** Full cost: new mapper + schema generation + parsing every time. */
    @Benchmark
    public void coldRead(Blackhole bh) throws IOException {
        SpreadsheetMapper cold = new SpreadsheetMapper();
        List<BenchRow> values = cold.readValues(file, BenchRow.class);
        bh.consume(values);
    }

    /** Reuse mapper (schema cached internally on first call). */
    @Benchmark
    public void warmMapper(Blackhole bh) throws IOException {
        List<BenchRow> values = mapper.readValues(file, BenchRow.class);
        bh.consume(values);
    }

    /** Reuse mapper + pre-built schema. Pure parsing cost. */
    @Benchmark
    public void warmSchema(Blackhole bh) throws IOException {
        @SuppressWarnings("unchecked")
        List<BenchRow> values = (List<BenchRow>) (List<?>) mapper.reader(schema).forType(BenchRow.class)
                .readValues(SheetInput.source(file)).readAll();
        bh.consume(values);
    }

    /** Schema generation cost only (no file I/O). */
    @Benchmark
    public void schemaOnly(Blackhole bh) throws Exception {
        SpreadsheetSchema s = mapper.sheetSchemaFor(BenchRow.class);
        bh.consume(s);
    }
}
