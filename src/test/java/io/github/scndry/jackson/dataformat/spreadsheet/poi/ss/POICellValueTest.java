package io.github.scndry.jackson.dataformat.spreadsheet.poi.ss;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class POICellValueTest {

    private Workbook _workbook;
    private Cell _cell;
    private CountingDataFormatter _formatter;

    @BeforeEach
    void setUp() {
        _workbook = new XSSFWorkbook();
        final Sheet sheet = _workbook.createSheet();
        final Row row = sheet.createRow(0);
        _cell = row.createCell(0);
        _cell.setCellValue(1234.56);
        final CellStyle style = _workbook.createCellStyle();
        style.setDataFormat(_workbook.createDataFormat().getFormat("\"$\"#,##0.00"));
        _cell.setCellStyle(style);
        _formatter = new CountingDataFormatter();
    }

    @AfterEach
    void tearDown() throws IOException {
        _workbook.close();
    }

    @Test
    void deferredUntilFirstGetStringValue() {
        final POICellValue value = new POICellValue(1234.56, _cell, _formatter);
        assertThat(_formatter.callCount).isZero();

        final String s = value.getStringValue();
        assertThat(_formatter.callCount).isEqualTo(1);
        assertThat(s).isEqualTo("$1,234.56");
    }

    @Test
    void memoizedAfterFirstCompute() {
        final POICellValue value = new POICellValue(1234.56, _cell, _formatter);
        value.getStringValue();
        value.getStringValue();
        value.getStringValue();
        assertThat(_formatter.callCount).isEqualTo(1);
    }

    @Test
    void numericPathSkipsCompute() {
        final POICellValue value = new POICellValue(1234.56, _cell, _formatter);
        assertThat(value.getNumberValue()).isEqualTo(1234.56);
        assertThat(_formatter.callCount).isZero();
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
