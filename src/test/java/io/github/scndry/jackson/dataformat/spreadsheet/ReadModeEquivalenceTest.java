package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Verifies that POI-mode and SSML-mode reads return the same String value
 * for NUMERIC cells mapped to a String field, across cell formats (currency,
 * percentage, date). Closes the cross-mode inconsistency targeted by #153.
 */
class ReadModeEquivalenceTest {

    private File _file;

    @Data
    @NoArgsConstructor
    @DataGrid
    static class FormattedRow {
        String price;
        String pct;
        String when;
    }

    @BeforeEach
    void setUp() throws IOException {
        _file = File.createTempFile("read-mode-equiv-", ".xlsx");
        _file.deleteOnExit();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            final Sheet sheet = wb.createSheet();

            final CellStyle currency = wb.createCellStyle();
            currency.setDataFormat(wb.createDataFormat().getFormat("\"$\"#,##0.00"));
            final CellStyle percentage = wb.createCellStyle();
            percentage.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
            final CellStyle date = wb.createCellStyle();
            date.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd"));

            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("price");
            header.createCell(1).setCellValue("pct");
            header.createCell(2).setCellValue("when");

            final Row data = sheet.createRow(1);
            final Cell priceCell = data.createCell(0);
            priceCell.setCellValue(1234.56);
            priceCell.setCellStyle(currency);
            final Cell pctCell = data.createCell(1);
            pctCell.setCellValue(0.5);
            pctCell.setCellStyle(percentage);
            final Cell dateCell = data.createCell(2);
            dateCell.setCellValue(45797.0); // Excel serial — used only for cross-mode equivalence
            dateCell.setCellStyle(date);

            try (FileOutputStream fos = new FileOutputStream(_file)) {
                wb.write(fos);
            }
        }
    }

    @AfterEach
    void tearDown() {
        _file.delete();
    }

    @Test
    void poiAndSsmlReturnSameStringFieldForFormattedNumericCells() throws IOException {
        final SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();
        final SpreadsheetMapper poiMapper = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));

        final FormattedRow ssmlRow = ssmlMapper.readValue(_file, FormattedRow.class);
        final FormattedRow poiRow = poiMapper.readValue(_file, FormattedRow.class);

        assertThat(ssmlRow.price).isEqualTo(poiRow.price);
        assertThat(ssmlRow.pct).isEqualTo(poiRow.pct);
        assertThat(ssmlRow.when).isEqualTo(poiRow.when);

        // Absolute checks for the deterministic formats; date formatting is
        // locale-sensitive so only the equivalence above is asserted for it.
        assertThat(poiRow.price).isEqualTo("$1,234.56");
        assertThat(poiRow.pct).isEqualTo("50.00%");
    }
}
