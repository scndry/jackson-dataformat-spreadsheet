package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for POI integration scenarios: template-based writing,
 * multi-sheet workbooks, and post-processing after data binding.
 */
class POIIntegrationTest {

    SpreadsheetMapper mapper;

    @TempDir
    Path tempDir;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Product {
        String name;
        int quantity;
        double price;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Order {
        @DataColumn("Order ID")
        int orderId;
        String customer;
        double total;
    }

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    @Test
    void templateBasedWrite() throws Exception {
        File template = tempFile("template.xlsx");
        File output = tempFile("output.xlsx");

        // Create a template with styled header row
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellStyle headerStyle = wb.createCellStyle();
            Font font = wb.createFont();
            font.setBold(true);
            headerStyle.setFont(font);

            Sheet sheet = wb.createSheet("Report");
            Row header = sheet.createRow(0);
            Cell h0 = header.createCell(0);
            h0.setCellValue("name");
            h0.setCellStyle(headerStyle);
            Cell h1 = header.createCell(1);
            h1.setCellValue("quantity");
            h1.setCellStyle(headerStyle);
            Cell h2 = header.createCell(2);
            h2.setCellValue("price");
            h2.setCellStyle(headerStyle);

            try (OutputStream os = new FileOutputStream(template)) {
                wb.write(os);
            }
        }

        // Open template, write data only (preserve header formatting)
        SpreadsheetMapper templateMapper = SpreadsheetMapper.builder()
                .origin("A2")
                .useHeader(false)
                .build();

        List<Product> products = Arrays.asList(
                new Product("Apple", 10, 1.50),
                new Product("Banana", 20, 0.80));

        try (XSSFWorkbook wb = new XSSFWorkbook(template)) {
            Sheet sheet = wb.getSheet("Report");
            templateMapper.writeValue(sheet, products, Product.class);
            try (OutputStream os = new FileOutputStream(output)) {
                wb.write(os);
            }
        }

        // Verify: data is written and template header formatting is preserved
        try (XSSFWorkbook wb = new XSSFWorkbook(output)) {
            Sheet sheet = wb.getSheet("Report");

            // Header row preserved with bold font
            Row header = sheet.getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(wb.getFontAt(header.getCell(0).getCellStyle().getFontIndex()).getBold())
                    .isTrue();

            // Data rows written correctly
            Row row1 = sheet.getRow(1);
            assertThat(row1.getCell(0).getStringCellValue()).isEqualTo("Apple");
            assertThat((int) row1.getCell(1).getNumericCellValue()).isEqualTo(10);
            assertThat(row1.getCell(2).getNumericCellValue()).isEqualTo(1.50);

            Row row2 = sheet.getRow(2);
            assertThat(row2.getCell(0).getStringCellValue()).isEqualTo("Banana");
        }
    }

    @Test
    void multiSheetWrite() throws Exception {
        File file = tempFile("multi-write.xlsx");

        List<Product> products = Arrays.asList(
                new Product("Apple", 10, 1.50),
                new Product("Banana", 20, 0.80));
        List<Order> orders = Arrays.asList(
                new Order(1, "Alice", 100.0),
                new Order(2, "Bob", 200.0));

        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            mapper.writeValue(wb.createSheet("Products"), products, Product.class);
            mapper.writeValue(wb.createSheet("Orders"), orders, Order.class);
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
            wb.dispose();
        }

        // Verify both sheets independently
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getNumberOfSheets()).isEqualTo(2);
            assertThat(wb.getSheet("Products")).isNotNull();
            assertThat(wb.getSheet("Orders")).isNotNull();
        }

        List<Product> readProducts = mapper.readValues(
                SheetInput.source(file, "Products"), Product.class);
        assertThat(readProducts).isEqualTo(products);

        List<Order> readOrders = mapper.readValues(
                SheetInput.source(file, "Orders"), Order.class);
        assertThat(readOrders).isEqualTo(orders);
    }

    @Test
    void multiSheetRead() throws Exception {
        File file = tempFile("multi-read.xlsx");

        // Prepare file with two sheets of different types
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            mapper.writeValue(wb.createSheet("Products"), Arrays.asList(
                    new Product("X", 5, 9.99)), Product.class);
            mapper.writeValue(wb.createSheet("Orders"), Arrays.asList(
                    new Order(99, "Charlie", 500.0)), Order.class);
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
            wb.dispose();
        }

        // Read both sheets from a single open workbook
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            List<Product> products = mapper.readValues(
                    wb.getSheet("Products"), Product.class);
            List<Order> orders = mapper.readValues(
                    wb.getSheet("Orders"), Order.class);

            assertThat(products).hasSize(1);
            assertThat(products.get(0).name).isEqualTo("X");

            assertThat(orders).hasSize(1);
            assertThat(orders.get(0).orderId).isEqualTo(99);
            assertThat(orders.get(0).customer).isEqualTo("Charlie");
        }
    }

    @Test
    void postProcessingWithFormula() throws Exception {
        File file = tempFile("formula.xlsx");

        List<Product> products = Arrays.asList(
                new Product("Apple", 10, 1.50),
                new Product("Banana", 20, 0.80),
                new Product("Cherry", 5, 3.00));

        // Data binding first, then POI for formula
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Sales");
            mapper.writeValue(sheet, products, Product.class);

            // Add SUM formula below the price column
            int lastDataRow = sheet.getLastRowNum();
            Row totalRow = sheet.createRow(lastDataRow + 1);
            totalRow.createCell(0).setCellValue("TOTAL");
            totalRow.createCell(2).setCellFormula(
                    "SUM(C2:C" + (lastDataRow + 1) + ")");

            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
            wb.dispose();
        }

        // Verify: data rows + formula row
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheet("Sales");

            // Data rows
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("Apple");
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("Cherry");

            // Formula row
            Row totalRow = sheet.getRow(4);
            assertThat(totalRow.getCell(0).getStringCellValue()).isEqualTo("TOTAL");
            assertThat(totalRow.getCell(2).getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(totalRow.getCell(2).getCellFormula()).isEqualTo("SUM(C2:C4)");
        }
    }

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }
}
