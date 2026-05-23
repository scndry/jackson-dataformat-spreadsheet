package io.github.scndry.jackson.dataformat.spreadsheet.internal;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetFactory;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

/**
 * Type-controlled cell-count-parity probe — flat 10 String columns vs
 * nested 6 String outer + 4 String inner. Mirrors
 * {@link NestedIntParityBenchmark}'s structure with {@code String}
 * everywhere so the shared-strings lookup path (SST cache hits) is the
 * dominant per-cell cost on the flat side. The remaining gap to the
 * nested path measures the record-tree algorithm's fixed overhead
 * against a cheaper-per-cell baseline than the int probe.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class NestedStringParityBenchmark {

    static final int ITEMS_PER_OUTER = 10;

    @Param({"100000", "1000000"})
    int targetCells;

    File flatFile;
    File nestedFile;
    List<StringRow> flatData;
    List<StringOuter> nestedData;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoiRead;
    SpreadsheetMapper mapperPoiWrite;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class StringRow {
        @DataColumn private String c0;
        @DataColumn private String c1;
        @DataColumn private String c2;
        @DataColumn private String c3;
        @DataColumn private String c4;
        @DataColumn private String c5;
        @DataColumn private String c6;
        @DataColumn private String c7;
        @DataColumn private String c8;
        @DataColumn private String c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    public static class StringItem {
        @DataColumn private String i0;
        @DataColumn private String i1;
        @DataColumn private String i2;
        @DataColumn private String i3;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class StringOuter {
        @DataColumn(anchor = true) private String id;
        @DataColumn private String o1;
        @DataColumn private String o2;
        @DataColumn private String o3;
        @DataColumn private String o4;
        @DataColumn private String o5;
        @DataColumnGroup("Items") private List<StringItem> items;
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        final int flatRows = targetCells / 10;
        final int nestedInner = Math.round(targetCells / 4.6f / ITEMS_PER_OUTER) * ITEMS_PER_OUTER;
        final int outerCount = nestedInner / ITEMS_PER_OUTER;

        mapper = new SpreadsheetMapper();
        mapperPoiRead = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapperPoiWrite = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));

        flatData = new ArrayList<>(flatRows);
        for (int i = 0; i < flatRows; i++) {
            flatData.add(new StringRow(
                    "c0-" + i, "c1-" + i, "c2-" + i, "c3-" + i, "c4-" + i,
                    "c5-" + i, "c6-" + i, "c7-" + i, "c8-" + i, "c9-" + i));
        }
        nestedData = new ArrayList<>(outerCount);
        for (int i = 0; i < outerCount; i++) {
            List<StringItem> items = new ArrayList<>(ITEMS_PER_OUTER);
            final int base = i * ITEMS_PER_OUTER;
            for (int j = 0; j < ITEMS_PER_OUTER; j++) {
                items.add(new StringItem(
                        "i0-" + (base + j), "i1-" + (base + j),
                        "i2-" + (base + j), "i3-" + (base + j)));
            }
            nestedData.add(new StringOuter(
                    "id-" + i, "o1-" + i, "o2-" + i, "o3-" + i, "o4-" + i, "o5-" + i, items));
        }

        flatFile = _writeFlatFile(flatData);
        nestedFile = _writeNestedFile(nestedData);
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        flatFile.delete();
        nestedFile.delete();
    }

    private File _writeFlatFile(final List<StringRow> data) throws IOException {
        File file = File.createTempFile("bench-string-parity-flat-", ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row header = sheet.createRow(0);
            for (int c = 0; c < 10; c++) header.createCell(c).setCellValue("c" + c);
            for (int i = 0; i < data.size(); i++) {
                StringRow r = data.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(r.getC0());
                row.createCell(1).setCellValue(r.getC1());
                row.createCell(2).setCellValue(r.getC2());
                row.createCell(3).setCellValue(r.getC3());
                row.createCell(4).setCellValue(r.getC4());
                row.createCell(5).setCellValue(r.getC5());
                row.createCell(6).setCellValue(r.getC6());
                row.createCell(7).setCellValue(r.getC7());
                row.createCell(8).setCellValue(r.getC8());
                row.createCell(9).setCellValue(r.getC9());
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        return file;
    }

    private File _writeNestedFile(final List<StringOuter> data) throws IOException {
        File file = File.createTempFile("bench-string-parity-nested-", ".xlsx");
        file.deleteOnExit();
        SpreadsheetMapper writer = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS));
        writer.writeValue(file, data, StringOuter.class);
        return file;
    }

    @Benchmark
    public void flatRead(Blackhole bh) throws IOException {
        List<StringRow> values = mapper.readValues(flatFile, StringRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flatReadPoi(Blackhole bh) throws IOException {
        List<StringRow> values = mapperPoiRead.readValues(flatFile, StringRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flatWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, flatData, StringRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void flatWritePoi(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapperPoiWrite.writeValue(out, flatData, StringRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void nestedRead(Blackhole bh) throws IOException {
        List<StringOuter> values = mapper.readValues(nestedFile, StringOuter.class);
        bh.consume(values);
    }

    @Benchmark
    public void nestedReadPoi(Blackhole bh) throws IOException {
        List<StringOuter> values = mapperPoiRead.readValues(nestedFile, StringOuter.class);
        bh.consume(values);
    }

    @Benchmark
    public void nestedWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, nestedData, StringOuter.class);
        bh.consume(out);
    }

    @Benchmark
    public void nestedWritePoi(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapperPoiWrite.writeValue(out, nestedData, StringOuter.class);
        bh.consume(out);
    }
}
