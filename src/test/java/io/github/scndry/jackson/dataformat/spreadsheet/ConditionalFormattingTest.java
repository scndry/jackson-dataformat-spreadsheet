package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.Formula;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.File;
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
                            .fillForegroundColor(IndexedColors.RED.getIndex())
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting()
                            .column("score")
                            .greaterThanOrEqual(80)
                            .style("highlight")
                            .end())
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
                            .fillForegroundColor(IndexedColors.RED.getIndex())
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting()
                            .column("score")
                            .greaterThanOrEqual(80)
                            .style("highlight")
                            .end())
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
                            .fillForegroundColor(IndexedColors.RED.getIndex())
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting()
                            .column("score")
                            .greaterThanOrEqual(80)
                            .style("highlight")
                            .end())
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
                    .fillForegroundColor(IndexedColors.RED.getIndex())
                    .fillPattern().solidForeground()
                    .end();
        GridConfigurer gridConfigurer = new GridConfigurer()
                .conditionalFormatting()
                    .column("score")
                    .greaterThanOrEqual(80)
                    .style("highlight")
                    .end();

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
                            .fillForegroundColor(IndexedColors.RED.getIndex())
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting()
                            .column("score")
                            .greaterThanOrEqual(80)
                            .style("highlight")
                            .end())
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
                spec -> spec.greaterThan(80L).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.GT);
        assertThat(rule.getFormula1()).isEqualTo("80");
    }

    @Test
    void typedNumericDouble() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.lessThan(80.5).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.LT);
        assertThat(rule.getFormula1()).isEqualTo("80.5");
    }

    @Test
    void typedBoolean() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.equalTo(true).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.EQUAL);
        assertThat(rule.getFormula1()).isEqualTo("TRUE");
    }

    @Test
    void typedStringAutoEscape() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("name",
                spec -> spec.equalTo("URGENT").style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.EQUAL);
        assertThat(rule.getFormula1()).isEqualTo("\"URGENT\"");
    }

    @Test
    void typedStringInternalQuoteEscape() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("name",
                spec -> spec.equalTo("say \"hi\"").style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        // Excel doubles internal quotes
        assertThat(rule.getFormula1()).isEqualTo("\"say \"\"hi\"\"\"");
    }

    @Test
    void typedLocalDate() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.greaterThan(LocalDate.of(2026, 1, 1)).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("DATE(2026,1,1)");
    }

    @Test
    void typedLocalDateTime() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.greaterThan(LocalDateTime.of(2026, 1, 1, 10, 30, 0)).style("highlight"),
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
                spec -> spec.greaterThan(date).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).startsWith("DATE(2026,1,1)+TIME(");
    }

    @Test
    void typedCalendar() throws Exception {
        Calendar cal = Calendar.getInstance(TimeZone.getDefault());
        cal.set(2026, Calendar.JANUARY, 1, 12, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.greaterThan(cal).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("DATE(2026,1,1)+TIME(12,0,0)");
    }

    @Test
    void typedBetween() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.between(80, 100).style("highlight"), Score.class,
                Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.BETWEEN);
        assertThat(rule.getFormula1()).isEqualTo("80");
        assertThat(rule.getFormula2()).isEqualTo("100");
    }

    @Test
    void formulaOfRawCellRef() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.greaterThan(Formula.of("$D$1")).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("$D$1");
    }

    @Test
    void formulaOfFunction() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("score",
                spec -> spec.greaterThan(Formula.of("AVERAGE($B$2:$B$100)")).style("highlight"),
                Score.class, Arrays.asList(new Score("Alice", 90)));
        assertThat(rule.getFormula1()).isEqualTo("AVERAGE($B$2:$B$100)");
    }

    @Test
    void formulaColumnResolvesToRowRelativeRef() throws Exception {
        ConditionalFormattingRule rule = _writeAndGetRule("price",
                spec -> spec.greaterThan(Formula.column("minPrice")).style("highlight"),
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
                            .fillForegroundColor(IndexedColors.RED.getIndex())
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting()
                            .column("price")
                            .greaterThan(Formula.column("nonexistent"))
                            .style("highlight")
                            .end())
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
                            .fillForegroundColor(IndexedColors.RED.getIndex())
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting()
                            .column("score")
                            .expression("AND($B2>0, $B2<100)")
                            .style("highlight")
                            .end())
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
        GridConfigurer configurer = new GridConfigurer();
        assertThatThrownBy(() -> configurer.conditionalFormatting()
                .column("score").expression("").style("highlight").end())
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    /** Helper: write a CF rule and return the resulting POI rule. */
    private <T> ConditionalFormattingRule _writeAndGetRule(
            String column,
            java.util.function.Function<io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormattingRuleSpec,
                io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormattingRuleSpec> ruleConfigurer,
            Class<T> type,
            List<T> data) throws Exception {
        File file = new File(tempDir, "cf-typed-" + System.nanoTime() + ".xlsx");
        GridConfigurer gc = new GridConfigurer();
        ruleConfigurer.apply(gc.conditionalFormatting().column(column)).end();
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("highlight")
                            .fillForegroundColor(IndexedColors.RED.getIndex())
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(gc)
                .build();
        mapper.writeValue(file, data, type);
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            return wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getConditionalFormattingAt(0).getRule(0);
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
