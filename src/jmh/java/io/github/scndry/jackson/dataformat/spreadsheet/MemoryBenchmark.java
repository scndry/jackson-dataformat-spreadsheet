package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.fesod.sheet.FesodSheet;
import com.poiji.bind.Poiji;
import com.poiji.option.PoijiOptions;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Memory-focused benchmark with realistic mixed-type data
 * and shared string table.
 * GC profiler is enabled by default in build.gradle.kts.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.SingleShotTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3)
@Measurement(iterations = 5)
@Fork(1)
public class MemoryBenchmark {

    @Param({"10000", "50000", "100000"})
    int rowCount;

    File file;

    private static final String[] CATEGORIES = {
            "Electronics", "Clothing", "Food", "Books", "Sports",
            "Home", "Garden", "Toys", "Health", "Automotive"
    };

    private static final String[] STATUSES = {
            "Active", "Inactive", "Pending", "Archived"
    };

    @Data @NoArgsConstructor @DataGrid
    public static class Product {
        public String name;
        public String category;
        public int quantity;
        public double price;
        public boolean inStock;
        public String status;
    }

    @Data @NoArgsConstructor
    public static class PoijiProduct {
        @com.poiji.annotation.ExcelCellName("name")
        public String name;
        @com.poiji.annotation.ExcelCellName("category")
        public String category;
        @com.poiji.annotation.ExcelCellName("quantity")
        public int quantity;
        @com.poiji.annotation.ExcelCellName("price")
        public double price;
        @com.poiji.annotation.ExcelCellName("inStock")
        public boolean inStock;
        @com.poiji.annotation.ExcelCellName("status")
        public String status;
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        file = File.createTempFile("bench-mem-", ".xlsx");
        file.deleteOnExit();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Products");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("name");
            header.createCell(1).setCellValue("category");
            header.createCell(2).setCellValue("quantity");
            header.createCell(3).setCellValue("price");
            header.createCell(4).setCellValue("inStock");
            header.createCell(5).setCellValue("status");
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("Product-" + i);
                row.createCell(1).setCellValue(CATEGORIES[i % CATEGORIES.length]);
                row.createCell(2).setCellValue(i % 500);
                row.createCell(3).setCellValue(9.99 + (i % 1000) * 0.5);
                row.createCell(4).setCellValue(i % 3 != 0);
                row.createCell(5).setCellValue(STATUSES[i % STATUSES.length]);
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void jacksonSpreadsheet(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        List<Product> values = mapper.readValues(file, Product.class);
        bh.consume(values);
    }

    @Benchmark
    public void poiUserModel(Blackhole bh) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<Product> entries = new ArrayList<>(rowCount);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Product p = new Product();
                p.name = row.getCell(0).getStringCellValue();
                p.category = row.getCell(1).getStringCellValue();
                p.quantity = (int) row.getCell(2).getNumericCellValue();
                p.price = row.getCell(3).getNumericCellValue();
                p.inStock = row.getCell(4).getBooleanCellValue();
                p.status = row.getCell(5).getStringCellValue();
                entries.add(p);
            }
            bh.consume(entries);
        }
    }

    @Benchmark
    public void poiji(Blackhole bh) {
        List<PoijiProduct> values = Poiji.fromExcel(
                file, PoijiProduct.class,
                PoijiOptions.PoijiOptionsBuilder.settings().build());
        bh.consume(values);
    }

    @Benchmark
    public void fesod(Blackhole bh) throws IOException {
        List<Product> values = FesodSheet.read(file)
                .head(Product.class).headRowNumber(1)
                .sheet().doReadSync();
        bh.consume(values);
    }

    @Benchmark
    public void fastExcelReader(Blackhole bh) throws IOException {
        try (ReadableWorkbook wb = new ReadableWorkbook(file)) {
            List<Product> entries = new ArrayList<>(rowCount);
            wb.getFirstSheet().openStream().forEach(row -> {
                if (row.getRowNum() == 1) return;
                Product p = new Product();
                p.name = row.getCellText(0);
                p.category = row.getCellText(1);
                p.quantity = (int) Double.parseDouble(
                        row.getCellText(2));
                p.price = Double.parseDouble(row.getCellText(3));
                p.inStock = Boolean.parseBoolean(
                        row.getCellText(4));
                p.status = row.getCellText(5);
                entries.add(p);
            });
            bh.consume(entries);
        }
    }
}
