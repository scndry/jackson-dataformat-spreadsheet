package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.*;

class StylesTest {

    @TempDir File tempDir;
    SpreadsheetMapper mapper;

    @DataGrid
    static class StyledEntry {
        @DataColumn(value = "Name", style = "nameStyle") public String name;
        @DataColumn(value = "Amount", style = "amountStyle") public int amount;
        @DataColumn(value = "Price", style = "priceStyle") public double price;
        public StyledEntry() {}
        public StyledEntry(String name, int amount, double price) {
            this.name = name; this.amount = amount; this.price = price;
        }
    }

    @BeforeEach
    void setUp() {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("nameStyle")
                    .dataFormat().text()
                    .end()
                .cellStyle("amountStyle")
                    .dataFormat().numberIntWithComma()
                    .alignment().right()
                    .end()
                .cellStyle("priceStyle")
                    .dataFormat().numberFloatWithComma()
                    .fillForegroundColor(0xFFFF00)
                    .fillPattern().solidForeground()
                    .font().bold().italic().fontHeight((short) 12).color(0xFF0000).end()
                    .end();

        mapper = SpreadsheetMapper.builder()
                .stylesBuilder(styles)
                .build();
    }

    @Test
    void stylesApplied() throws Exception {
        File file = new File(tempDir, "styled.xlsx");
        List<StyledEntry> data = Arrays.asList(
                new StyledEntry("Alice", 1000, 19.99),
                new StyledEntry("Bob", 2000, 29.99));
        mapper.writeValue(file, data, StyledEntry.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(1);
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo("Alice");
            assertThat(row.getCell(1).getNumericCellValue()).isEqualTo(1000.0);
            assertThat(row.getCell(2).getNumericCellValue()).isEqualTo(19.99);
            // Verify styles exist
            assertThat(row.getCell(1).getCellStyle().getDataFormatString()).contains("#,##0");
            assertThat(row.getCell(2).getCellStyle().getDataFormatString()).contains("#,##0.00");
        }
    }

    @Test
    void cloneStyleFrom() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("base")
                    .dataFormat().numberInt()
                    .border().thin()
                    .end()
                .cellStyle("derived", "base")
                    .fillForegroundColor(0x00FF00)
                    .fillPattern().solidForeground()
                    .end();

        SpreadsheetMapper m = SpreadsheetMapper.builder().stylesBuilder(styles).build();
        File file = new File(tempDir, "cloned.xlsx");

        @DataGrid class Row { @DataColumn(style = "derived") public int value = 42; }

        m.writeValue(file, new Row());
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertThat(style.getBorderLeft()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getFillPattern()).isEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    @Test
    void cloneStyleFromUnregisteredThrows() {
        assertThatThrownBy(() -> new StylesBuilder().cellStyle("x", "nonexistent"))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("nonexistent");
    }

    @Test
    void unregisteredStyleNameThrows() throws Exception {
        @DataGrid class Bad { @DataColumn(style = "missing") public int value = 1; }

        File file = new File(tempDir, "bad.xlsx");
        assertThatThrownBy(() -> mapper.writeValue(file, new Bad()))
                .hasRootCauseInstanceOf(IllegalStateException.class)
                .hasRootCauseMessage("Style 'missing' is not registered");
    }

    @Test
    void simpleFactoryAppliesTypeDefaults() throws Exception {
        StylesBuilder styles = StylesBuilder.simple();
        SpreadsheetMapper m = SpreadsheetMapper.builder().stylesBuilder(styles).build();

        @DataGrid class TypedRow {
            public int primitiveInt = 1000;
            public double primitiveDouble = 1.5;
            public Integer wrappedInt = 2000;
            public String text = "hello";
        }

        File file = new File(tempDir, "simple.xlsx");
        m.writeValue(file, new TypedRow());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row row = wb.getSheetAt(0).getRow(1);
            // primitive int → comma format
            assertThat(row.getCell(0).getCellStyle().getDataFormatString()).contains("#,##0");
            // primitive double → comma float format
            assertThat(row.getCell(1).getCellStyle().getDataFormatString()).contains("#,##0.00");
            // wrapped Integer → plain int
            assertThat(row.getCell(2).getCellStyle().getDataFormatString()).isEqualTo("0");
            // String → text format
            assertThat(row.getCell(3).getCellStyle().getDataFormatString()).isEqualTo("@");
        }
    }

    @Test
    void fontProperties() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("fancy")
                    .font()
                        .fontName("Arial")
                        .fontHeight((short) 14)
                        .bold()
                        .italic()
                        .strikeout()
                        .underline().single()
                    .end()
                .end();

        SpreadsheetMapper m = SpreadsheetMapper.builder().stylesBuilder(styles).build();

        @DataGrid class Row { @DataColumn(style = "fancy") public String value = "test"; }

        File file = new File(tempDir, "font.xlsx");
        m.writeValue(file, new Row());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertThat(font.getFontName()).isEqualTo("Arial");
            assertThat(font.getFontHeightInPoints()).isEqualTo((short) 14);
            assertThat(font.getBold()).isTrue();
            assertThat(font.getItalic()).isTrue();
            assertThat(font.getStrikeout()).isTrue();
            assertThat(font.getUnderline()).isEqualTo(FontUnderline.SINGLE.getByteValue());
        }
    }

    @Test
    void cellStyleProperties() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("full")
                    .alignment().center()
                    .verticalAlignment().top()
                    .wrapText()
                    .rotation((short) 45)
                    .indention((short) 2)
                    .shrinkToFit()
                    .hidden()
                    .locked()
                    .border().medium()
                    .borderColor(0x0000FF)
                    .fillForegroundColor(IndexedColors.YELLOW)
                    .fillPattern().solidForeground()
                    .dataFormat("0.00%")
                    .end();

        SpreadsheetMapper m = SpreadsheetMapper.builder().stylesBuilder(styles).build();

        @DataGrid class Row { @DataColumn(style = "full") public double value = 0.5; }

        File file = new File(tempDir, "full.xlsx");
        m.writeValue(file, new Row());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertThat(style.getAlignment()).isEqualTo(HorizontalAlignment.CENTER);
            assertThat(style.getVerticalAlignment()).isEqualTo(VerticalAlignment.TOP);
            assertThat(style.getWrapText()).isTrue();
            assertThat(style.getRotation()).isEqualTo((short) 45);
            assertThat(style.getIndention()).isEqualTo((short) 2);
            assertThat(style.getShrinkToFit()).isTrue();
            assertThat(style.getHidden()).isTrue();
            assertThat(style.getLocked()).isTrue();
            assertThat(style.getBorderLeft()).isEqualTo(BorderStyle.MEDIUM);
            assertThat(style.getBorderRight()).isEqualTo(BorderStyle.MEDIUM);
            assertThat(style.getFillPattern()).isEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(style.getDataFormatString()).isEqualTo("0.00%");
        }
    }

    @Test
    void headerStyle() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("hdr")
                    .font().bold().end()
                    .end();

        @DataGrid(columnHeaderStyle = "hdr")
        class Row { public int a = 1; public int b = 2; }

        SpreadsheetMapper m = SpreadsheetMapper.builder().stylesBuilder(styles).build();
        File file = new File(tempDir, "header.xlsx");
        m.writeValue(file, new Row());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row0: {
                org.apache.poi.ss.usermodel.Row headerRow = wb.getSheetAt(0).getRow(0);
                Font font = wb.getFontAt(headerRow.getCell(0).getCellStyle().getFontIndex());
                assertThat(font.getBold()).isTrue();
            }
        }
    }

    @Test
    void borderDirectSetters() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("mixed")
                    .borderLeft().thin()
                    .borderRight().thick()
                    .borderTop().dashed()
                    .borderBottom().dotted()
                    .leftBorderColor(0xFF0000)
                    .rightBorderColor(0x00FF00)
                    .topBorderColor(0x0000FF)
                    .bottomBorderColor(0xFFFF00)
                    .end();

        SpreadsheetMapper m = SpreadsheetMapper.builder().stylesBuilder(styles).build();

        @DataGrid class Row { @DataColumn(style = "mixed") public int v = 1; }

        File file = new File(tempDir, "borders.xlsx");
        m.writeValue(file, new Row());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertThat(style.getBorderLeft()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getBorderRight()).isEqualTo(BorderStyle.THICK);
            assertThat(style.getBorderTop()).isEqualTo(BorderStyle.DASHED);
            assertThat(style.getBorderBottom()).isEqualTo(BorderStyle.DOTTED);
        }
    }

    @Test
    void colorByRgbComponents() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("rgb")
                    .fillForegroundColor(255, 128, 0)
                    .fillPattern().solidForeground()
                    .end();

        SpreadsheetMapper m = SpreadsheetMapper.builder().stylesBuilder(styles).build();

        @DataGrid class Row { @DataColumn(style = "rgb") public int v = 1; }

        File file = new File(tempDir, "rgb.xlsx");
        m.writeValue(file, new Row());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertThat(style.getFillPattern()).isEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    @Test
    void existingCellStylePreservedWhenNoStyleConfigured() throws Exception {
        // Create a file with pre-styled cells
        File file = new File(tempDir, "pre-styled.xlsx");
        short bgColorIndex;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellStyle styledCell = wb.createCellStyle();
            styledCell.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            styledCell.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            bgColorIndex = styledCell.getIndex();

            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("name");
            header.createCell(1).setCellValue("value");
            Row data = sheet.createRow(1);
            Cell c0 = data.createCell(0);
            c0.setCellStyle(styledCell);
            Cell c1 = data.createCell(1);
            c1.setCellStyle(styledCell);
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        // Write data with no styles configured — existing styles should be preserved
        @DataGrid class Simple {
            public String name = "Alice";
            public int value = 42;
        }

        File output = new File(tempDir, "pre-styled-output.xlsx");
        SpreadsheetMapper noStyleMapper = SpreadsheetMapper.builder()
                .origin("A2")
                .useHeader(false)
                .build();
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            noStyleMapper.writeValue(sheet, new Simple());
            try (OutputStream os = new FileOutputStream(output)) {
                wb.write(os);
            }
        }

        // Verify: data written AND style preserved
        try (XSSFWorkbook wb = new XSSFWorkbook(output)) {
            Sheet sheet = wb.getSheetAt(0);
            Cell c0 = sheet.getRow(1).getCell(0);
            assertThat(c0.getStringCellValue()).isEqualTo("Alice");
            assertThat(c0.getCellStyle().getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(c0.getCellStyle().getFillForegroundColor())
                    .isEqualTo(IndexedColors.YELLOW.getIndex());
        }
    }
}
