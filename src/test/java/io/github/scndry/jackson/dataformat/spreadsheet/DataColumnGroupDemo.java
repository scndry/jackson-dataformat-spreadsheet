package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfSystemProperty;

import io.github.scndry.jackson.dataformat.spreadsheet.DataColumnGroupTest.Address;
import io.github.scndry.jackson.dataformat.spreadsheet.DataColumnGroupTest.Company;
import io.github.scndry.jackson.dataformat.spreadsheet.DataColumnGroupTest.Employee;
import io.github.scndry.jackson.dataformat.spreadsheet.DataColumnGroupTest.QuarterMetrics;
import io.github.scndry.jackson.dataformat.spreadsheet.DataColumnGroupTest.YearMetrics;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;

/**
 * Visual demo — writes group output files to {@code build/data-column-group-demo/}
 * so LibreOffice / Excel can open them directly. Disabled by default;
 * enable with {@code ./gradlew test -Pdemo=true}.
 */
@EnabledIfSystemProperty(named = "demo", matches = "true")
class DataColumnGroupDemo {

    private static final Path OUTPUT_DIR = Paths.get("build", "data-column-group-demo");

    @Test
    void writeDemoFiles() throws Exception {
        SpreadsheetMapper poiMapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();

        Path outDir = OUTPUT_DIR;
        Files.createDirectories(outDir);

        Employee employee = new Employee(1, "Alice", new Address("Seoul", "12345"));
        Company company = new Company("Acme",
                new YearMetrics(new QuarterMetrics(100, 20), new QuarterMetrics(110, 22)),
                new YearMetrics(new QuarterMetrics(120, 25), new QuarterMetrics(130, 28)));

        File poi1 = outDir.resolve("group-1depth-poi.xlsx").toFile();
        poiMapper.writeValue(poi1, employee);
        File poi2 = outDir.resolve("group-2depth-poi.xlsx").toFile();
        poiMapper.writeValue(poi2, company);

        File ssml1 = outDir.resolve("group-1depth-ssml.xlsx").toFile();
        ssmlMapper.writeValue(ssml1, employee);
        File ssml2 = outDir.resolve("group-2depth-ssml.xlsx").toFile();
        ssmlMapper.writeValue(ssml2, company);

        System.out.println("PoC demo files written to: " + outDir.toAbsolutePath());
        System.out.println("  - " + poi1.getAbsolutePath());
        System.out.println("  - " + poi2.getAbsolutePath());
        System.out.println("  - " + ssml1.getAbsolutePath());
        System.out.println("  - " + ssml2.getAbsolutePath());
    }

    /**
     * Visual side-effect check — multi-row header combined with origin offset,
     * autoFilter, and freezePane. Confirms that:
     *   - origin "B2" shifts everything down/right (group rows + leaf + data)
     *   - autoFilter arrows land on the leaf header row (not the top group row)
     *   - freezePane covers all header rows + everything above origin
     */
    @Test
    void writeDemoWithFeatures() throws Exception {
        // origin = B2  → originRow=1
        // 1-depth header → headerRowCount=2 → leafRow=2, dataRow=3
        // freezePane rowSplit=3 → freeze rows 0..2 (everything above data)
        GridConfigurer grid = new GridConfigurer()
                .freezePane(0, 3)
                .autoFilter();

        SpreadsheetMapper poiMapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL))
                .setOrigin("B2")
                .setGridConfigurer(grid);
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper()
                .setOrigin("B2")
                .setGridConfigurer(grid);

        Path outDir = OUTPUT_DIR;
        Files.createDirectories(outDir);

        List<Employee> values = Arrays.asList(
                new Employee(1, "Alice", new Address("Seoul", "12345")),
                new Employee(2, "Bob", new Address("Busan", "23456")),
                new Employee(3, "Carol", new Address("Incheon", "34567")),
                new Employee(4, "Dave", new Address("Daegu", "45678")),
                new Employee(5, "Eve", new Address("Daejeon", "56789")));

        File poi = outDir.resolve("group-with-features-poi.xlsx").toFile();
        poiMapper.writeValue(poi, values, Employee.class);

        File ssml = outDir.resolve("group-with-features-ssml.xlsx").toFile();
        ssmlMapper.writeValue(ssml, values, Employee.class);

        System.out.println("Features demo files:");
        System.out.println("  - " + poi.getAbsolutePath());
        System.out.println("  - " + ssml.getAbsolutePath());
    }

    /**
     * Control case — same origin / autoFilter / freezePane setup but no
     * {@link io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup}.
     * Useful for diffing visual side effects against {@link #writeDemoWithFeatures()}.
     */
    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class FlatEmployee {
        @DataColumn("ID") int id;
        @DataColumn("Name") String name;
        @DataColumn("City") String city;
        @DataColumn("Zip") String zip;
    }

    @Test
    void writeNonGroupDemoWithFeatures() throws Exception {
        // origin = B2  → originRow=1
        // No group → headerRowCount=1 → leafRow=1, dataRow=2
        // freezePane rowSplit=2 (dataRow) → freeze rows 0..1, equivalent semantic to
        // the group case (freeze everything above data).
        GridConfigurer grid = new GridConfigurer()
                .freezePane(0, 2)
                .autoFilter();

        SpreadsheetMapper poiMapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL))
                .setOrigin("B2")
                .setGridConfigurer(grid);
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper()
                .setOrigin("B2")
                .setGridConfigurer(grid);

        Path outDir = OUTPUT_DIR;
        Files.createDirectories(outDir);

        List<FlatEmployee> values = Arrays.asList(
                new FlatEmployee(1, "Alice", "Seoul", "12345"),
                new FlatEmployee(2, "Bob", "Busan", "23456"),
                new FlatEmployee(3, "Carol", "Incheon", "34567"),
                new FlatEmployee(4, "Dave", "Daegu", "45678"),
                new FlatEmployee(5, "Eve", "Daejeon", "56789"));

        File poi = outDir.resolve("non-group-with-features-poi.xlsx").toFile();
        poiMapper.writeValue(poi, values, FlatEmployee.class);

        File ssml = outDir.resolve("non-group-with-features-ssml.xlsx").toFile();
        ssmlMapper.writeValue(ssml, values, FlatEmployee.class);

        System.out.println("Non-group features demo files:");
        System.out.println("  - " + poi.getAbsolutePath());
        System.out.println("  - " + ssml.getAbsolutePath());
    }
}
