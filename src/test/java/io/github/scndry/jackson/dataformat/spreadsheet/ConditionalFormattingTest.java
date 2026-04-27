package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.feature.ConditionalFormattingConfigurer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.File;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

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
                .conditionalFormattings(new ConditionalFormattingConfigurer()
                        .rule()
                            .column("score")
                            .greaterThanOrEqual("80")
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
                .conditionalFormattings(new ConditionalFormattingConfigurer()
                        .rule()
                            .column("score")
                            .greaterThanOrEqual("80")
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
                .conditionalFormattings(new ConditionalFormattingConfigurer()
                        .rule()
                            .column("score")
                            .greaterThanOrEqual("80")
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
        ConditionalFormattingConfigurer conditionalFormattings = new ConditionalFormattingConfigurer()
                .rule()
                    .column("score")
                    .greaterThanOrEqual("80")
                    .style("highlight")
                    .end();

        // SSML path (default)
        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder()
                .stylesBuilder(styles)
                .conditionalFormattings(conditionalFormattings)
                .build();
        ssmlMapper.writeValue(ssmlFile, data, Score.class);

        // POI path
        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .stylesBuilder(styles)
                .conditionalFormattings(conditionalFormattings)
                .build();
        poiMapper.writeValue(poiFile, data, Score.class);

        // Compare dxf entries in styles.xml
        _assertDxfEquivalent(poiFile, ssmlFile);

        // Compare cfRule in sheet1.xml (ignoring sqref range differences)
        _assertCfRuleEquivalent(poiFile, ssmlFile);
    }

    private static void _assertDxfEquivalent(File expected, File actual) throws Exception {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(true);
        String ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        try (OPCPackage expPkg = OPCPackage.open(expected);
             OPCPackage actPkg = OPCPackage.open(actual)) {
            Document expDoc = _parsePart(expPkg, "/xl/styles.xml", dbf);
            Document actDoc = _parsePart(actPkg, "/xl/styles.xml", dbf);

            NodeList expDxfs = expDoc.getElementsByTagNameNS(ns, "dxf");
            NodeList actDxfs = actDoc.getElementsByTagNameNS(ns, "dxf");
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
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(true);
        String ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        try (OPCPackage expPkg = OPCPackage.open(expected);
             OPCPackage actPkg = OPCPackage.open(actual)) {
            Document expDoc = _parsePart(expPkg, "/xl/worksheets/sheet1.xml", dbf);
            Document actDoc = _parsePart(actPkg, "/xl/worksheets/sheet1.xml", dbf);

            NodeList expRules = expDoc.getElementsByTagNameNS(ns, "cfRule");
            NodeList actRules = actDoc.getElementsByTagNameNS(ns, "cfRule");
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
                NodeList expFormulas = expRule.getElementsByTagNameNS(ns, "formula");
                NodeList actFormulas = actRule.getElementsByTagNameNS(ns, "formula");
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

    private static Document _parsePart(OPCPackage pkg, String partName,
            DocumentBuilderFactory dbf) throws Exception {
        try (InputStream is = pkg.getPart(
                PackagingURIHelper.createPartName(partName)).getInputStream()) {
            return dbf.newDocumentBuilder().parse(is);
        }
    }
}
