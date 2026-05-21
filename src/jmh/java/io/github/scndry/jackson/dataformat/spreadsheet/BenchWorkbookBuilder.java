package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public final class BenchWorkbookBuilder {

    private BenchWorkbookBuilder() {}

    public static File createSampleFile(final String prefix, final int rowCount) throws IOException {
        File file = File.createTempFile(prefix, ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row header = sheet.createRow(0);
            for (int c = 0; c < BenchRow.HEADERS.length; c++) {
                header.createCell(c).setCellValue(BenchRow.HEADERS[c]);
            }
            for (int i = 0; i < rowCount; i++) {
                BenchRow r = BenchRow.create(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(r.getId());
                row.createCell(1).setCellValue(r.getName());
                row.createCell(2).setCellValue(r.getCategory());
                row.createCell(3).setCellValue(r.getStatus());
                row.createCell(4).setCellValue(r.getQuantity());
                row.createCell(5).setCellValue(r.getPrice());
                row.createCell(6).setCellValue(r.getAmount().doubleValue());
                row.createCell(7).setCellValue(r.getDueDate());
                row.createCell(8).setCellValue(r.getDescription());
                row.createCell(9).setCellValue(r.getCreatedAt());
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        return file;
    }
}
