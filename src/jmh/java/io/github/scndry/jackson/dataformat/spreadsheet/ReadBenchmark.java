package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.fesod.sheet.FesodSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

/**
 * Realistic read benchmark: 10-column schema covering identity (Long),
 * business data (int, double, BigDecimal, LocalDate, String × 4), and an
 * audit timestamp (LocalDateTime).
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class ReadBenchmark {

    @Param({"1000", "10000", "50000", "100000", "500000"})
    int rowCount;

    File file;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoi;

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        mapper = new SpreadsheetMapper();
        mapperPoi = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        file = BenchWorkbookBuilder.createSampleFile("bench-read-", rowCount);
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void jacksonSpreadsheet(Blackhole bh) throws IOException {
        List<BenchRow> values = mapper.readValues(file, BenchRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void jacksonSpreadsheetPOI(Blackhole bh) throws IOException {
        List<BenchRow> values = mapperPoi.readValues(file, BenchRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void poiUserModel(Blackhole bh) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<BenchRow> entries = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                BenchRow r = new BenchRow();
                r.setId((long) row.getCell(0).getNumericCellValue());
                r.setName(row.getCell(1).getStringCellValue());
                r.setCategory(row.getCell(2).getStringCellValue());
                r.setStatus(row.getCell(3).getStringCellValue());
                r.setQuantity((int) row.getCell(4).getNumericCellValue());
                r.setPrice(row.getCell(5).getNumericCellValue());
                r.setAmount(BigDecimal.valueOf(row.getCell(6).getNumericCellValue()));
                r.setDueDate(row.getCell(7).getLocalDateTimeCellValue().toLocalDate());
                r.setDescription(row.getCell(8).getStringCellValue());
                r.setCreatedAt(row.getCell(9).getLocalDateTimeCellValue());
                entries.add(r);
            }
            bh.consume(entries);
        }
    }

    @Benchmark
    public void fesod(Blackhole bh) throws IOException {
        List<BenchRow> values = FesodSheet.read(file)
                .head(BenchRow.class).headRowNumber(1)
                .sheet().doReadSync();
        bh.consume(values);
    }

    @Benchmark
    public void fastExcelReader(Blackhole bh) throws IOException {
        try (ReadableWorkbook wb = new ReadableWorkbook(file)) {
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
}
