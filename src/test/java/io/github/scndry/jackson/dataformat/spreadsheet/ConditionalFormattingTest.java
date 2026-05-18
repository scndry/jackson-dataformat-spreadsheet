package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormatRule;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.NS_SPREADSHEETML;
import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.parsePart;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.between;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.colorScale;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.columnRef;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.equalTo;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.expression;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.formula;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.greaterThan;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.greaterThanOrEqual;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.lessThan;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.lessThanOrEqual;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.notBetween;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.notEqualTo;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ConditionalFormattingTest {

    @TempDir File tempDir;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Score {
        private String name;
        private int score;
    }

    @Test
    void poiPathConditionalFormatting() throws Exception {
        File file = new File(tempDir, "cf-poi.xlsx");
        List<Score> data = Arrays.asList(
                new Score("Alice", 90),
                new Score("Bob", 50),
                new Score("Carol", 85));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                greaterThanOrEqual(80).style("highlight")))
                .build();

        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertThat(scf.getNumConditionalFormattings()).isEqualTo(1);
            ConditionalFormatting cf = scf.getConditionalFormattingAt(0);
            assertThat(cf.getNumberOfRules()).isEqualTo(1);
            ConditionalFormattingRule rule = cf.getRule(0);
            assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.GE);
            assertThat(rule.getFormula1()).isEqualTo("80");
            assertThat(rule.getPatternFormatting()).isNotNull();
        }
    }

    @Test
    void ssmlPathConditionalFormatting() throws Exception {
        File file = new File(tempDir, "cf-ssml.xlsx");
        List<Score> data = Arrays.asList(
                new Score("Alice", 90),
                new Score("Bob", 50),
                new Score("Carol", 85));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                greaterThanOrEqual(80).style("highlight")))
                .build();

        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertThat(scf.getNumConditionalFormattings()).isEqualTo(1);
            ConditionalFormattingRule rule = scf.getConditionalFormattingAt(0).getRule(0);
            assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.GE);
            assertThat(rule.getFormula1()).isEqualTo("80");
            assertThat(rule.getPatternFormatting()).isNotNull();
        }
    }

    @Test
    void noFeaturesConfigured() throws Exception {
        File file = new File(tempDir, "cf-none.xlsx");
        List<Score> data = Arrays.asList(new Score("Alice", 90));

        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertThat(scf.getNumConditionalFormattings()).isEqualTo(0);
        }
    }

    @Test
    void dataIntegrityWithConditionalFormatting() throws Exception {
        File file = new File(tempDir, "cf-roundtrip.xlsx");
        List<Score> data = Arrays.asList(
                new Score("Alice", 90),
                new Score("Bob", 50));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                greaterThanOrEqual(80).style("highlight")))
                .build();

        mapper.writeValue(file, data, Score.class);
        List<Score> read = new SpreadsheetMapper().readValues(file, Score.class);
        assertThat(read).isEqualTo(data);
    }

    @Test
    void domEquivalence() throws Exception {
        File ssmlFile = new File(tempDir, "cf-dom-ssml.xlsx");
        File poiFile = new File(tempDir, "cf-dom-poi.xlsx");
        List<Score> data = Arrays.asList(
                new Score("Alice", 90),
                new Score("Bob", 50));

        StylesBuilder styles = new StylesBuilder()
                .cellStyle("highlight")
                    .fillForegroundColor(IndexedColors.RED)
                    .fillPattern().solidForeground()
                    .end();
        GridConfigurer gridConfigurer = new GridConfigurer()
                .conditionalFormatting("score",
                        greaterThanOrEqual(80).style("highlight"));

        // SSML path (default)
        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder()
                .stylesBuilder(styles)
                .gridConfigurer(gridConfigurer)
                .build();
        ssmlMapper.writeValue(ssmlFile, data, Score.class);

        // POI path
        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(styles)
                .gridConfigurer(gridConfigurer)
                .build();
        poiMapper.writeValue(poiFile, data, Score.class);

        // Compare dxf entries in styles.xml
        _assertDxfEquivalent(poiFile, ssmlFile);

        // Compare cfRule in sheet1.xml (ignoring sqref range differences)
        _assertCfRuleEquivalent(poiFile, ssmlFile);
    }

    @Test
    void poiPathZeroDataRows() throws Exception {
        File file = new File(tempDir, "cf-poi-zero-rows.xlsx");
        List<Score> data = Collections.emptyList();

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                greaterThanOrEqual(80).style("highlight")))
                .build();

        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            // No data rows — CF rules are skipped
            assertThat(scf.getNumConditionalFormattings()).isZero();
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Item {
        private String name;
        private double price;
        private double minPrice;
    }

    @Test
    void typedNumericLong() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                greaterThan(80L).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.GT);
        assertThat(rule.getFormula1()).isEqualTo("80");
    }

    @Test
    void typedNumericDouble() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                lessThan(80.5).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.LT);
        assertThat(rule.getFormula1()).isEqualTo("80.5");
    }

    @Test
    void typedBoolean() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                equalTo(true).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.EQUAL);
        assertThat(rule.getFormula1()).isEqualTo("TRUE");
    }

    @Test
    void typedStringAutoEscape() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("name",
                equalTo("URGENT").style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.EQUAL);
        assertThat(rule.getFormula1()).isEqualTo("\"URGENT\"");
    }

    @Test
    void typedStringInternalQuoteEscape() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("name",
                equalTo("say \"hi\"").style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        // Excel doubles internal quotes
        assertThat(rule.getFormula1()).isEqualTo("\"say \"\"hi\"\"\"");
    }

    @Test
    void typedLocalDate() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                greaterThan(LocalDate.of(2026, 1, 1)).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("DATE(2026,1,1)");
    }

    @Test
    void typedLocalDateTime() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                greaterThan(LocalDateTime.of(2026, 1, 1, 10, 30, 0)).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("DATE(2026,1,1)+TIME(10,30,0)");
    }

    @Test
    void typedDate() throws Exception {
        Calendar cal = Calendar.getInstance(TimeZone.getDefault());
        cal.set(2026, Calendar.JANUARY, 1, 12, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        Date date = cal.getTime();
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                greaterThan(date).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).startsWith("DATE(2026,1,1)+TIME(");
    }

    @Test
    void typedCalendar() throws Exception {
        Calendar cal = Calendar.getInstance(TimeZone.getDefault());
        cal.set(2026, Calendar.JANUARY, 1, 12, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                greaterThan(cal).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("DATE(2026,1,1)+TIME(12,0,0)");
    }

    @Test
    void typedBetween() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                between(80, 100).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.BETWEEN);
        assertThat(rule.getFormula1()).isEqualTo("80");
        assertThat(rule.getFormula2()).isEqualTo("100");
    }

    @Test
    void lessThanOrEqualMapsToLE() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                lessThanOrEqual(20).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.LE);
        assertThat(rule.getFormula1()).isEqualTo("20");
    }

    @Test
    void notEqualToMapsToNotEqual() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                notEqualTo(0).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.NOT_EQUAL);
        assertThat(rule.getFormula1()).isEqualTo("0");
    }

    @Test
    void notBetweenMapsToNotBetween() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                notBetween(0, 20).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.NOT_BETWEEN);
        assertThat(rule.getFormula1()).isEqualTo("0");
        assertThat(rule.getFormula2()).isEqualTo("20");
    }

    @Test
    void formulaOfRawCellRef() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                greaterThan(formula("$D$1")).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("$D$1");
    }

    @Test
    void formulaOfFunction() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                greaterThan(formula("AVERAGE($B$2:$B$100)")).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("AVERAGE($B$2:$B$100)");
    }

    @Test
    void formulaColumnResolvesToRowRelativeRef() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("price",
                greaterThan(columnRef("minPrice")).style("highlight"),
                Item.class, Arrays.asList(new Item("Apple", 10.0, 5.0)));
        // Item: name=A, price=B, minPrice=C; data starts at row 2 (origin A1, header at row 1)
        assertThat(rule.getFormula1()).isEqualTo("$C2");
    }

    @Test
    void formulaColumnMissingColumnFails() throws Exception {
        File file = new File(tempDir, "cf-missing-col.xlsx");
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("price",
                                greaterThan(columnRef("nonexistent")).style("highlight")))
                .build();
        assertThatThrownBy(() -> mapper.writeValue(file, Arrays.asList(new Item("A", 1.0, 0.5)), Item.class))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("Formula column 'nonexistent' not found");
    }

    @Test
    void expressionRule() throws Exception {
        File file = new File(tempDir, "cf-expression.xlsx");
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                expression("AND($B2>0, $B2<100)").style("highlight")))
                .build();
        mapper.writeValue(file, Arrays.asList(new Score("Alice", 90)), Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            ConditionalFormattingRule rule = wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getConditionalFormattingAt(0).getRule(0);
            assertThat(rule.getConditionType()).isEqualTo(ConditionType.FORMULA);
            assertThat(rule.getFormula1()).isEqualTo("AND($B2>0, $B2<100)");
        }
    }

    @Test
    void expressionRejectsEmpty() {
        assertThatThrownBy(() -> expression(""))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void multipleRulesOnSameColumn() throws Exception {
        File file = new File(tempDir, "cf-multi.xlsx");
        List<Score> data = Arrays.asList(new Score("Alice", 90), new Score("Bob", 50));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("good").fillForegroundColor(IndexedColors.LIGHT_GREEN)
                            .fillPattern().solidForeground().end()
                        .cellStyle("bad").fillForegroundColor(IndexedColors.ROSE)
                            .fillPattern().solidForeground().end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                greaterThanOrEqual(80).style("good"),
                                lessThan(60).style("bad")))
                .build();
        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertThat(scf.getNumConditionalFormattings()).isEqualTo(2);
            assertThat(scf.getConditionalFormattingAt(0).getRule(0).getComparisonOperation())
                    .isEqualTo(ComparisonOperator.GE);
            assertThat(scf.getConditionalFormattingAt(1).getRule(0).getComparisonOperation())
                    .isEqualTo(ComparisonOperator.LT);
        }
    }

    @Test
    void colorScaleDefaults() throws Exception {
        File file = new File(tempDir, "cf-cs-default.xlsx");
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score", colorScale()))
                .build();
        mapper.writeValue(file, Arrays.asList(new Score("Alice", 90)), Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            ConditionalFormattingRule rule = wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getConditionalFormattingAt(0).getRule(0);
            assertThat(rule.getConditionType()).isEqualTo(ConditionType.COLOR_SCALE);
            ConditionalFormattingThreshold[] t = rule.getColorScaleFormatting().getThresholds();
            assertThat(t).hasSize(3);
            assertThat(t[0].getRangeType()).isEqualTo(RangeType.MIN);
            assertThat(t[1].getRangeType()).isEqualTo(RangeType.PERCENTILE);
            assertThat(t[1].getValue()).isEqualTo(50.0);
            assertThat(t[2].getRangeType()).isEqualTo(RangeType.MAX);
            assertThat(rule.getColorScaleFormatting().getColors()).hasSize(3);
        }
    }

    @Test
    void colorScaleExplicitThresholds() throws Exception {
        File file = new File(tempDir, "cf-cs-explicit.xlsx");
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score", colorScale(0, 50, 100)))
                .build();
        mapper.writeValue(file, Arrays.asList(new Score("Alice", 90)), Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            ConditionalFormattingRule rule = wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getConditionalFormattingAt(0).getRule(0);
            assertThat(rule.getConditionType()).isEqualTo(ConditionType.COLOR_SCALE);
            ConditionalFormattingThreshold[] t = rule.getColorScaleFormatting().getThresholds();
            assertThat(t[0].getRangeType()).isEqualTo(RangeType.NUMBER);
            assertThat(t[0].getValue()).isEqualTo(0.0);
            assertThat(t[1].getRangeType()).isEqualTo(RangeType.NUMBER);
            assertThat(t[1].getValue()).isEqualTo(50.0);
            assertThat(t[2].getRangeType()).isEqualTo(RangeType.NUMBER);
            assertThat(t[2].getValue()).isEqualTo(100.0);
        }
    }

    @Test
    void unknownColumnThrows() {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("x").fillForegroundColor(IndexedColors.YELLOW)
                            .fillPattern().solidForeground().end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("nonexistent", greaterThanOrEqual(80).style("x")))
                .build();
        File file = new File(tempDir, "cf-fail-col.xlsx");
        assertThatThrownBy(() -> mapper.writeValue(file, Arrays.asList(new Score("A", 1)), Score.class))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("nonexistent");
    }

    @Test
    void unknownStyleThrows() {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score", greaterThanOrEqual(80).style("missing-style")))
                .build();
        File file = new File(tempDir, "cf-fail-style.xlsx");
        assertThatThrownBy(() -> mapper.writeValue(file, Arrays.asList(new Score("A", 1)), Score.class))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("missing-style");
    }

    @Test
    void styleWithNullNameThrows() {
        assertThatThrownBy(() -> greaterThanOrEqual(80).style(null))
                .isInstanceOf(NullPointerException.class)
                .hasMessageContaining("style name");
    }

    @Test
    void conditionalFormattingNullColumnThrows() {
        assertThatThrownBy(() -> new GridConfigurer()
                .conditionalFormatting(null, greaterThanOrEqual(80).style("x")))
                .isInstanceOf(NullPointerException.class)
                .hasMessageContaining("column");
    }

    @Test
    void conditionalFormattingNullRuleElementThrows() {
        assertThatThrownBy(() -> new GridConfigurer()
                .conditionalFormatting("score",
                        greaterThanOrEqual(80).style("x"),
                        (ConditionalFormatRule) null))
                .isInstanceOf(NullPointerException.class)
                .hasMessageContaining("rules element");
    }

    @Test
    void formulaWithNullTextThrows() {
        assertThatThrownBy(() -> formula(null))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void formulaWithEmptyTextThrows() {
        assertThatThrownBy(() -> formula(""))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void columnRefWithNullNameThrows() {
        assertThatThrownBy(() -> columnRef(null))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void columnRefWithEmptyNameThrows() {
        assertThatThrownBy(() -> columnRef(""))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    /** Helper: write a CF rule and return the resulting POI rule. */
    private <T> ConditionalFormattingRule _writeAndGetRule(
            String column,
            ConditionalFormatRule rule,
            Class<T> type,
            List<T> data) throws Exception {
        File file = new File(tempDir, "cf-typed-" + System.nanoTime() + ".xlsx");
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting(column, rule))
                .build();
        mapper.writeValue(file, data, type);
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            return wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getConditionalFormattingAt(0).getRule(0);
        }
    }

    @Test
    void hssfPathConditionalFormattingFontAndBorder() throws Exception {
        // HSSF (.xls) workbook routes font color, underline, font height, and
        // border color through the indexed-color path of GridConfigurer's
        // dxf builder — distinct from the XSSF color object path.
        File file = new File(tempDir, "cf-hssf.xls");
        List<Score> data = Arrays.asList(new Score("Alice", 90), new Score("Bob", 50));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .font()
                                .bold()
                                .fontHeight((short) 14)
                                .color(IndexedColors.WHITE)
                                .underline().single()
                                .end()
                            .borderTop().thin()
                            .topBorderColor(IndexedColors.RED)
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                greaterThanOrEqual(80).style("highlight")))
                .build();

        try (HSSFWorkbook wb = new HSSFWorkbook();
             OutputStream os = new FileOutputStream(file)) {
            Sheet sheet = wb.createSheet("Data");
            mapper.writeValue(sheet, data, Score.class);
            wb.write(os);
        }

        try (InputStream is = new FileInputStream(file);
             HSSFWorkbook wb = new HSSFWorkbook(is)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertThat(scf.getNumConditionalFormattings()).isEqualTo(1);
            ConditionalFormattingRule rule = scf.getConditionalFormattingAt(0).getRule(0);
            assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.GE);
            assertThat(rule.getFormula1()).isEqualTo("80");

            // Indexed-path font dxf: color index, underline byte, height in twentieths of a point.
            FontFormatting ff = rule.getFontFormatting();
            assertThat(ff).isNotNull();
            assertThat(ff.isBold()).isTrue();
            assertThat(ff.getFontColorIndex()).isEqualTo(IndexedColors.WHITE.getIndex());
            assertThat(ff.getUnderlineType()).isEqualTo(Font.U_SINGLE);
            assertThat(ff.getFontHeight()).isEqualTo(14 * 20);

            // Indexed-path border dxf: top color resolved to indexed value.
            BorderFormatting bf = rule.getBorderFormatting();
            assertThat(bf).isNotNull();
            assertThat(bf.getBorderTop()).isEqualTo(BorderStyle.THIN);
            assertThat(bf.getTopBorderColor()).isEqualTo(IndexedColors.RED.getIndex());
        }
    }

    private static void _assertDxfEquivalent(File expected, File actual) throws Exception {
        try (OPCPackage expPkg = OPCPackage.open(expected);
             OPCPackage actPkg = OPCPackage.open(actual)) {
            Document expDoc = parsePart(expPkg, "/xl/styles.xml");
            Document actDoc = parsePart(actPkg, "/xl/styles.xml");

            NodeList expDxfs = expDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "dxf");
            NodeList actDxfs = actDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "dxf");
            assertThat(actDxfs.getLength())
                    .as("dxf count")
                    .isEqualTo(expDxfs.getLength());
            for (int i = 0; i < expDxfs.getLength(); i++) {
                assertThat(actDxfs.item(i).isEqualNode(expDxfs.item(i)))
                        .as("dxf[%d] DOM equality", i)
                        .isTrue();
            }
        }
    }

    private static void _assertCfRuleEquivalent(File expected, File actual) throws Exception {
        try (OPCPackage expPkg = OPCPackage.open(expected);
             OPCPackage actPkg = OPCPackage.open(actual)) {
            Document expDoc = parsePart(expPkg, "/xl/worksheets/sheet1.xml");
            Document actDoc = parsePart(actPkg, "/xl/worksheets/sheet1.xml");

            NodeList expRules = expDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "cfRule");
            NodeList actRules = actDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "cfRule");
            assertThat(actRules.getLength())
                    .as("cfRule count")
                    .isEqualTo(expRules.getLength());
            for (int i = 0; i < expRules.getLength(); i++) {
                Element expRule = (Element) expRules.item(i);
                Element actRule = (Element) actRules.item(i);
                // Compare attributes (type, operator, dxfId, priority)
                assertThat(actRule.getAttribute("type"))
                        .as("cfRule[%d] type", i)
                        .isEqualTo(expRule.getAttribute("type"));
                assertThat(actRule.getAttribute("operator"))
                        .as("cfRule[%d] operator", i)
                        .isEqualTo(expRule.getAttribute("operator"));
                assertThat(actRule.getAttribute("dxfId"))
                        .as("cfRule[%d] dxfId", i)
                        .isEqualTo(expRule.getAttribute("dxfId"));
                // Compare formula content
                NodeList expFormulas = expRule.getElementsByTagNameNS(NS_SPREADSHEETML, "formula");
                NodeList actFormulas = actRule.getElementsByTagNameNS(NS_SPREADSHEETML, "formula");
                assertThat(actFormulas.getLength())
                        .as("cfRule[%d] formula count", i)
                        .isEqualTo(expFormulas.getLength());
                for (int j = 0; j < expFormulas.getLength(); j++) {
                    assertThat(actFormulas.item(j).getTextContent())
                            .as("cfRule[%d] formula[%d]", i, j)
                            .isEqualTo(expFormulas.item(j).getTextContent());
                }
            }
        }
    }

}
