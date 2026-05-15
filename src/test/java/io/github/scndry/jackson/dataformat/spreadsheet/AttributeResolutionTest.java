package io.github.scndry.jackson.dataformat.spreadsheet;

import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.Data;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Attribute resolution order is cross-cutting across {@link DataColumn},
 * {@link DataColumnGroup}, and {@link DataGrid}. Each leaf column
 * attribute resolves in priority order:
 *
 * <ol>
 *   <li>{@code @DataColumn} on the property</li>
 *   <li>innermost enclosing {@code @DataColumnGroup}</li>
 *   <li>outer enclosing {@code @DataColumnGroup} (recurse outward)</li>
 *   <li>{@code @DataGrid} on the declaring class</li>
 *   <li>{@code @DataGrid} on the root class</li>
 * </ol>
 *
 * First-non-default value wins. Tests below pin the order at the
 * Value layer (no schema) and at the {@link SpreadsheetSchema} layer.
 */
class AttributeResolutionTest {

    // ----------------------------------------------------------------
    // Value layer — withDefaults without going through schema build
    // ----------------------------------------------------------------

    private static DataColumnGroup.Value group(final String name, final String headerStyle) {
        return new DataColumnGroup.Value(name, "", headerStyle,
                "", "", DataGrid.DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT,
                DataGrid.DEFAULT_MIN_COLUMN_WIDTH, DataGrid.DEFAULT_MAX_COLUMN_WIDTH,
                OptBoolean.DEFAULT);
    }

    private static DataColumnGroup.Value groupWithChildDefaults(
            final String columnStyle, final String columnHeaderStyle,
            final OptBoolean merge) {
        return new DataColumnGroup.Value("g", "", "",
                columnStyle, columnHeaderStyle,
                DataGrid.DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT,
                DataGrid.DEFAULT_MIN_COLUMN_WIDTH, DataGrid.DEFAULT_MAX_COLUMN_WIDTH,
                merge);
    }

    @Test
    void groupOwnHeaderStyle_explicitWinsOverDataGrid() {
        final DataGrid.Value grid = new DataGrid.Value("", "", "gridGroupBg",
                DataGrid.DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT,
                DataGrid.DEFAULT_MIN_COLUMN_WIDTH, DataGrid.DEFAULT_MAX_COLUMN_WIDTH,
                OptBoolean.DEFAULT);
        final DataColumnGroup.Value g = group("Year", "annotationGroupBg").withDefaults(grid);
        assertThat(g.getHeaderStyle()).isEqualTo("annotationGroupBg");
    }

    @Test
    void groupOwnHeaderStyle_inheritsDataGridWhenUnset() {
        final DataGrid.Value grid = new DataGrid.Value("", "", "gridGroupBg",
                DataGrid.DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT,
                DataGrid.DEFAULT_MIN_COLUMN_WIDTH, DataGrid.DEFAULT_MAX_COLUMN_WIDTH,
                OptBoolean.DEFAULT);
        final DataColumnGroup.Value g = group("Year", "").withDefaults(grid);
        assertThat(g.getHeaderStyle()).isEqualTo("gridGroupBg");
    }

    @Test
    void childHeaderStyle_leafWinsOverGroup() {
        final DataColumn.Value leaf = DataColumn.Value.from(null)
                .withDefaults(groupWithChildDefaults("", "groupChildHeader",
                        OptBoolean.DEFAULT).asChildDefaults());
        assertThat(leaf.getHeaderStyle()).isEqualTo("groupChildHeader");
    }

    @Test
    void childHeaderStyle_innermostGroupWinsOverOuter() {
        final DataColumn.Value leaf = DataColumn.Value.from(null)
                .withDefaults(groupWithChildDefaults("", "innerChild", OptBoolean.DEFAULT)
                        .asChildDefaults())
                .withDefaults(groupWithChildDefaults("", "outerChild", OptBoolean.DEFAULT)
                        .asChildDefaults());
        assertThat(leaf.getHeaderStyle()).isEqualTo("innerChild");
    }

    @Test
    void childHeaderStyle_fallsThroughToDataGridWhenGroupUnset() {
        final DataGrid.Value grid = new DataGrid.Value("", "gridHeader", "",
                DataGrid.DEFAULT_COLUMN_WIDTH, OptBoolean.DEFAULT,
                DataGrid.DEFAULT_MIN_COLUMN_WIDTH, DataGrid.DEFAULT_MAX_COLUMN_WIDTH,
                OptBoolean.DEFAULT);
        final DataColumn.Value leaf = DataColumn.Value.from(null)
                .withDefaults(DataColumnGroup.Value.empty().asChildDefaults())
                .withDefaults(grid);
        assertThat(leaf.getHeaderStyle()).isEqualTo("gridHeader");
    }

    @Test
    void childMerge_cascadesFromGroup() {
        final DataColumn.Value leaf = DataColumn.Value.from(null)
                .withDefaults(groupWithChildDefaults("", "", OptBoolean.TRUE).asChildDefaults());
        assertThat(leaf.isMerge()).isTrue();
    }

    // ----------------------------------------------------------------
    // Schema layer — full cascade through Outer / Group / Element.
    //
    //   Outer  @DataGrid(columnStyle=…, autoSizeColumn=TRUE)
    //     ├ a                                                  (outside group)
    //     ├ @DataColumnGroup(columnHeaderStyle=…) List<Group>
    //     │    ├ b
    //     │    ├ @DataColumnGroup(columnStyle=…) List<Element>
    //     │    │    └ c  @DataColumn(style=…)
    //     │    └ d
    //     └ e                                                  (outside group)
    // ----------------------------------------------------------------

    @Data
    @DataGrid(columnStyle = "Outer_columnStyle", autoSizeColumn = OptBoolean.TRUE)
    static class Outer {
        int a;
        @DataColumnGroup(value = "groups", columnHeaderStyle = "Outer_groups_columnHeaderStyle")
        List<Group> groups;
        int e;
    }

    @Data
    @DataGrid(columnHeaderStyle = "Group_columnHeaderStyle")
    static class Group {
        int b;
        @DataColumnGroup(value = "elements", columnStyle = "Group_elements_columnStyle")
        List<Element> elements;
        int d;
    }

    @Data
    @DataGrid(autoSizeColumn = OptBoolean.FALSE)
    static class Element {
        @DataColumn(style = "Element_c_style") int c;
    }

    private static SpreadsheetSchema _schemaFor(final Class<?> type) throws Exception {
        return new SpreadsheetMapper().sheetSchemaFor(type);
    }

    private static Column _byName(final SpreadsheetSchema schema, final String name) {
        for (final Column c : schema) {
            if (c != null && name.equals(c.getName())) return c;
        }
        throw new AssertionError("No column named " + name);
    }

    @Test
    void cascade_outerLeaf_outsideGroup_resolvesFromRootGrid() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Outer.class);
        final Column a = _byName(schema, "a");
        // style: leaf/groups not in chain → Outer.grid wins
        assertThat(a.getValue().getStyle()).isEqualTo("Outer_columnStyle");
        assertThat(a.getValue().getHeaderStyle()).isEmpty();
        assertThat(a.getValue().getAutoSize()).isEqualTo(OptBoolean.TRUE);

        final Column e = _byName(schema, "e");
        assertThat(e.getValue().getStyle()).isEqualTo("Outer_columnStyle");
        assertThat(e.getValue().getHeaderStyle()).isEmpty();
        assertThat(e.getValue().getAutoSize()).isEqualTo(OptBoolean.TRUE);
    }

    @Test
    void cascade_middleLeaf_resolvesFromGroupAndRootGrid() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Outer.class);
        final Column b = _byName(schema, "groups/[]/b");
        // style: leaf / "groups" / Group.grid all unset → Outer.grid wins
        assertThat(b.getValue().getStyle()).isEqualTo("Outer_columnStyle");
        // headerStyle: innermost "groups" wins — Group.grid never reached
        assertThat(b.getValue().getHeaderStyle()).isEqualTo("Outer_groups_columnHeaderStyle");
        // autoSize: all unset until Outer.grid → TRUE
        assertThat(b.getValue().getAutoSize()).isEqualTo(OptBoolean.TRUE);

        final Column d = _byName(schema, "groups/[]/d");
        assertThat(d.getValue().getStyle()).isEqualTo("Outer_columnStyle");
        assertThat(d.getValue().getHeaderStyle()).isEqualTo("Outer_groups_columnHeaderStyle");
        assertThat(d.getValue().getAutoSize()).isEqualTo(OptBoolean.TRUE);
    }

    @Test
    void cascade_innermostLeaf_walksAllLevels() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Outer.class);
        final Column c = _byName(schema, "groups/[]/elements/[]/c");
        // style: leaf wins — "elements" / Element.grid / Outer.grid never reached
        assertThat(c.getValue().getStyle()).isEqualTo("Element_c_style");
        // headerStyle: "elements" unset → outer "groups" wins
        assertThat(c.getValue().getHeaderStyle()).isEqualTo("Outer_groups_columnHeaderStyle");
        // autoSize: leaf / "elements" / "groups" all unset → Element.grid (declaring) FALSE wins,
        // Outer.grid TRUE never reached.
        assertThat(c.getValue().getAutoSize()).isEqualTo(OptBoolean.FALSE);
    }
}
