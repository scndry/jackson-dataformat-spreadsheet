package io.github.scndry.jackson.dataformat.spreadsheet;

import com.alibaba.excel.EasyExcel;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Realistic read benchmark: mixed types, shared string table (XSSFWorkbook).
 * Simulates data exported from Excel/Google Sheets.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class ReadBenchmark {

    @Param({"1000", "10000", "50000", "100000"})
    int rowCount;

    File file;

    private static final String[] CATEGORIES = {
            "Electronics", "Clothing", "Food", "Books", "Sports",
            "Home", "Garden", "Toys", "Health", "Automotive"
    };

    private static final String[] STATUSES = {"Active", "Inactive", "Pending", "Archived"};

    @DataGrid
    public static class Product {
        public String name;
        public String category;
        public int quantity;
        public double price;
        public boolean inStock;
        public String status;

        public Product() {}
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        file = File.createTempFile("bench-read-", ".xlsx");
        file.deleteOnExit();
        // XSSFWorkbook to ensure shared string table (like real Excel files)
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
                bh.consume(p);
            }
        }
    }

    @Benchmark
    public void easyExcel(Blackhole bh) throws IOException {
        List<Product> values = EasyExcel.read(file).head(Product.class).headRowNumber(1).sheet().doReadSync();
        bh.consume(values);
    }

    @Benchmark
    public void fastExcelReader(Blackhole bh) throws IOException {
        try (ReadableWorkbook wb = new ReadableWorkbook(file)) {
            wb.getFirstSheet().openStream().forEach(row -> {
                if (row.getRowNum() == 1) return;
                Product p = new Product();
                p.name = row.getCellText(0);
                p.category = row.getCellText(1);
                p.quantity = (int) Double.parseDouble(row.getCellText(2));
                p.price = Double.parseDouble(row.getCellText(3));
                p.inStock = Boolean.parseBoolean(row.getCellText(4));
                p.status = row.getCellText(5);
                bh.consume(p);
            });
        }
    }
}
