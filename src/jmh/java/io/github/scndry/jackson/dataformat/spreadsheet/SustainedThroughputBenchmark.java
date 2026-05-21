package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

/**
 * Sustained-throughput benchmark over the shared 10-column schema. Runs each
 * library in a 60-second measurement window to surface steady-state cost
 * (allocation rate, GC pause accumulation) that single-shot per-op figures
 * can hide.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.Throughput)
@OutputTimeUnit(TimeUnit.SECONDS)
@Warmup(iterations = 1, time = 10)
@Measurement(iterations = 1, time = 60)
@Fork(1)
public class SustainedThroughputBenchmark {

    static final int ROW_COUNT = 10_000;
    File readSource;
    List<BenchRow> writeSource;
    SpreadsheetMapper mapper;

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        mapper = new SpreadsheetMapper();
        readSource = BenchWorkbookBuilder.createSampleFile("bench-sustained-", ROW_COUNT);
        writeSource = new ArrayList<>();
        for (int i = 0; i < ROW_COUNT; i++) {
            writeSource.add(BenchRow.create(i));
        }
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        readSource.delete();
    }

    @Benchmark
    public void readJackson(Blackhole bh) throws IOException {
        List<BenchRow> values = mapper.readValues(readSource, BenchRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void readFastExcel(Blackhole bh) throws IOException {
        try (ReadableWorkbook wb = new ReadableWorkbook(readSource)) {
            List<BenchRow> entries = new ArrayList<>();
            wb.getFirstSheet().openStream().forEach(row -> {
                if (row.getRowNum() == 1) return;
                BenchRow r = new BenchRow();
                r.setId(row.getCellAsNumber(0).orElseThrow().longValue());
                r.setName(row.getCellText(1));
                r.setCategory(row.getCellText(2));
                r.setStatus(row.getCellText(3));
                r.setQuantity(row.getCellAsNumber(4).orElseThrow().intValue());
                r.setPrice(row.getCellAsNumber(5).orElseThrow().doubleValue());
                r.setAmount(row.getCellAsNumber(6).orElseThrow());
                r.setDueDate(row.getCellAsDate(7).orElseThrow().toLocalDate());
                r.setDescription(row.getCellText(8));
                r.setCreatedAt(row.getCellAsDate(9).orElseThrow());
                entries.add(r);
            });
            bh.consume(entries);
        }
    }

    @Benchmark
    public void writeJackson(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, writeSource, BenchRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void writeFastExcel(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (Workbook wb = new Workbook(out, "bench", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            for (int c = 0; c < BenchRow.HEADERS.length; c++) {
                ws.value(0, c, BenchRow.HEADERS[c]);
            }
            for (int i = 0; i < writeSource.size(); i++) {
                BenchRow r = writeSource.get(i);
                int row = i + 1;
                ws.value(row, 0, r.getId());
                ws.value(row, 1, r.getName());
                ws.value(row, 2, r.getCategory());
                ws.value(row, 3, r.getStatus());
                ws.value(row, 4, r.getQuantity());
                ws.value(row, 5, r.getPrice());
                ws.value(row, 6, r.getAmount());
                ws.value(row, 7, r.getDueDate().atStartOfDay());
                ws.value(row, 8, r.getDescription());
                ws.value(row, 9, r.getCreatedAt());
            }
        }
        bh.consume(out);
    }
}
