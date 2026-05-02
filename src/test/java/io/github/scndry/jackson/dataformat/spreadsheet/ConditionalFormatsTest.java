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

import java.io.File;
import java.util.Arrays;
import java.util.List;

import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.colorScale;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.expression;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.greaterThanOrEqual;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.lessThan;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Smoke tests for the 1.6.0 flat conditional formatting API
 * ({@code conditionalFormatting(String, ConditionalFormatRule, ConditionalFormatRule...)}).
 */
class ConditionalFormatsTest {

    @TempDir File tempDir;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Score {
        private String name;
        private int score;
    }

    @Test
    void singleCellIsRule() throws Exception {
        File file = new File(tempDir, "single.xlsx");
        List<Score> data = Arrays.asList(new Score("Alice", 90), new Score("Bob", 50));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("good")
                            .fillForegroundColor(IndexedColors.LIGHT_GREEN)
                            .fillPattern().solidForeground()
                            .end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score", greaterThanOrEqual(80).style("good")))
                .build();
        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertThat(scf.getNumConditionalFormattings()).isEqualTo(1);
            ConditionalFormattingRule rule = scf.getConditionalFormattingAt(0).getRule(0);
            assertThat(rule.getConditionType()).isEqualTo(ConditionType.CELL_VALUE_IS);
            assertThat(rule.getComparisonOperation()).isEqualTo(ComparisonOperator.GE);
            assertThat(rule.getFormula1()).isEqualTo("80");
            assertThat(rule.getPatternFormatting()).isNotNull();
        }
    }

    @Test
    void multipleRulesOnSameColumn() throws Exception {
        File file = new File(tempDir, "multi.xlsx");
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
    void expressionRule() throws Exception {
        File file = new File(tempDir, "expr.xlsx");
        List<Score> data = Arrays.asList(new Score("Alice", 90));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("hi").fillForegroundColor(IndexedColors.YELLOW)
                            .fillPattern().solidForeground().end())
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score",
                                expression("AND($B2>0, $A2<>\"\")").style("hi")))
                .build();
        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            ConditionalFormattingRule rule = scf.getConditionalFormattingAt(0).getRule(0);
            assertThat(rule.getConditionType()).isEqualTo(ConditionType.FORMULA);
            assertThat(rule.getFormula1()).isEqualTo("AND($B2>0, $A2<>\"\")");
        }
    }

    @Test
    void colorScaleDefaults() throws Exception {
        File file = new File(tempDir, "cs-default.xlsx");
        List<Score> data = Arrays.asList(new Score("Alice", 90));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score", colorScale()))
                .build();
        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            ConditionalFormattingRule rule = scf.getConditionalFormattingAt(0).getRule(0);
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
        File file = new File(tempDir, "cs-explicit.xlsx");
        List<Score> data = Arrays.asList(new Score("Alice", 90));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer()
                        .conditionalFormatting("score", colorScale(0, 50, 100)))
                .build();
        mapper.writeValue(file, data, Score.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            ConditionalFormattingRule rule = scf.getConditionalFormattingAt(0).getRule(0);
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
        File file = new File(tempDir, "fail.xlsx");
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
        File file = new File(tempDir, "fail2.xlsx");
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
}
