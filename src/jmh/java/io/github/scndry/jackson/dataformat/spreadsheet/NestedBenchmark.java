package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Measures ColumnPointer caching effect on nested POJO (depth 3).
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class NestedBenchmark {

    @Param({"10000", "50000"})
    int rowCount;

    File file;

    @DataGrid
    public static class Order {
        public int id;
        public Customer customer;
        public Payment payment;

        public Order() {}
    }

    public static class Customer {
        public String name;
        public Address address;

        public Customer() {}
    }

    public static class Address {
        public String city;
        public String zip;

        public Address() {}
    }

    public static class Payment {
        public double amount;
        public String method;

        public Payment() {}
    }

    // Columns: id | customer/name | customer/address/city | customer/address/zip | payment/amount | payment/method

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        file = File.createTempFile("bench-nested-", ".xlsx");
        file.deleteOnExit();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("id");
            header.createCell(1).setCellValue("name");
            header.createCell(2).setCellValue("city");
            header.createCell(3).setCellValue("zip");
            header.createCell(4).setCellValue("amount");
            header.createCell(5).setCellValue("method");
            String[] cities = {"Seoul", "Busan", "Daegu", "Incheon", "Daejeon"};
            String[] methods = {"Card", "Cash", "Transfer", "Mobile"};
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(i);
                row.createCell(1).setCellValue("Customer-" + i);
                row.createCell(2).setCellValue(cities[i % cities.length]);
                row.createCell(3).setCellValue("0" + (10000 + i % 90000));
                row.createCell(4).setCellValue(9.99 + (i % 1000) * 0.5);
                row.createCell(5).setCellValue(methods[i % methods.length]);
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
    public void nestedRead(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        List<Order> values = mapper.readValues(file, Order.class);
        bh.consume(values);
    }
}
