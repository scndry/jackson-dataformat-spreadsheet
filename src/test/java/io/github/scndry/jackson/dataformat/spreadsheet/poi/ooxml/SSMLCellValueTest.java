package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class SSMLCellValueTest {

    private File _file;
    private SSMLWorkbook _workbook;
    private StylesLookup _styles;
    private int _styleIndex;
    private CountingDataFormatter _formatter;

    @BeforeEach
    void setUp() throws IOException {
        _file = File.createTempFile("ssml-cell-value-test-", ".xlsx");
        _file.deleteOnExit();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            final CellStyle style = wb.createCellStyle();
            style.setDataFormat(wb.createDataFormat().getFormat("\"$\"#,##0.00"));
            _styleIndex = style.getIndex();
            final Sheet sheet = wb.createSheet();
            final Row row = sheet.createRow(0);
            final Cell cell = row.createCell(0);
            cell.setCellValue(1234.56);
            cell.setCellStyle(style);
            try (FileOutputStream fos = new FileOutputStream(_file)) {
                wb.write(fos);
            }
        }
        _workbook = SSMLWorkbook.create(_file);
        _styles = new LazyStylesLookup(_workbook.getStylesPart());
        _formatter = new CountingDataFormatter();
    }

    @AfterEach
    void tearDown() throws IOException {
        _workbook.close();
        _file.delete();
    }

    @Test
    void deferredUntilFirstGetStringValue() {
        final SSMLCellValue value = new SSMLCellValue(
                1234.56, "1234.56", _styleIndex, _styles, _formatter);
        assertThat(_formatter.callCount).isZero();

        final String s = value.getStringValue();
        assertThat(_formatter.callCount).isEqualTo(1);
        assertThat(s).isEqualTo("$1,234.56");
    }

    @Test
    void memoizedAfterFirstCompute() {
        final SSMLCellValue value = new SSMLCellValue(
                1234.56, "1234.56", _styleIndex, _styles, _formatter);
        value.getStringValue();
        value.getStringValue();
        value.getStringValue();
        assertThat(_formatter.callCount).isEqualTo(1);
    }

    @Test
    void numericPathSkipsCompute() {
        final SSMLCellValue value = new SSMLCellValue(
                1234.56, "1234.56", _styleIndex, _styles, _formatter);
        assertThat(value.getNumberValue()).isEqualTo(1234.56);
        assertThat(_formatter.callCount).isZero();
    }

    @Test
    void outOfBoundsStyleIndexFallsBackToGeneralFormat() {
        final SSMLCellValue value = new SSMLCellValue(
                1234.56, "1234.56", Integer.MAX_VALUE, _styles, _formatter);
        // out-of-bounds → numFmtId 0 (General) → raw-equivalent output via DataFormatter
        assertThat(value.getStringValue()).isEqualTo("1234.56");
        assertThat(_formatter.callCount).isEqualTo(1);
    }

    private static final class CountingDataFormatter extends DataFormatter {
        int callCount;

        @Override
        public String formatRawCellContents(final double value, final int formatIndex, final String formatString) {
            callCount++;
            return super.formatRawCellContents(value, formatIndex, formatString);
        }
    }
}
