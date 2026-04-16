package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import com.fasterxml.jackson.annotation.OptBoolean;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.*;

class AnnotationValueTest {

    SpreadsheetMapper mapper = new SpreadsheetMapper();

    @Test
    void dataColumnDefaults() throws Exception {
        @DataGrid class Row {
            @DataColumn("Name") public String name;
            @DataColumn(value = "Qty", width = 20) public int qty;
        }

        SpreadsheetSchema schema = mapper.sheetSchemaFor(Row.class);
        List<Column> columns = schema.getColumns(ColumnPointer.empty());
        assertThat(columns).hasSize(2);
        assertThat(columns.get(0).getName()).isEqualTo("Name");
        assertThat(columns.get(1).getName()).isEqualTo("Qty");
    }

    @Test
    void dataGridMergeColumn() throws Exception {
        @DataGrid(mergeColumn = OptBoolean.TRUE)
        class Row {
            @DataColumn("A") public int a;
            @DataColumn("B") public int b;
        }

        SpreadsheetSchema schema = mapper.sheetSchemaFor(Row.class);
        for (Column col : schema.getColumns(ColumnPointer.empty())) {
            assertThat(col.isMerge()).isTrue();
        }
    }

    @Test
    void dataColumnMergeOverridesGridDefault() throws Exception {
        @DataGrid(mergeColumn = OptBoolean.TRUE)
        class Row {
            @DataColumn(value = "A", merge = OptBoolean.FALSE) public int a;
            @DataColumn("B") public int b;
        }

        SpreadsheetSchema schema = mapper.sheetSchemaFor(Row.class);
        assertThat(schema.getColumns(ColumnPointer.empty()).get(0).isMerge()).isFalse();
        assertThat(schema.getColumns(ColumnPointer.empty()).get(1).isMerge()).isTrue();
    }

    @Test
    void dataColumnAutoSize() throws Exception {
        @DataGrid(autoSizeColumn = OptBoolean.TRUE)
        class Row {
            @DataColumn(value = "A", autoSize = OptBoolean.FALSE) public int a;
            @DataColumn("B") public int b;
        }

        SpreadsheetSchema schema = mapper.sheetSchemaFor(Row.class);
        assertThat(schema.getColumns(ColumnPointer.empty()).get(0).getValue().isAutoSize()).isFalse();
        assertThat(schema.getColumns(ColumnPointer.empty()).get(1).getValue().isAutoSize()).isTrue();
    }

    @Test
    void dataColumnWithName() {
        DataColumn.Value val = new DataColumn.Value("original", "", "", 0, OptBoolean.DEFAULT, 0, 0, OptBoolean.DEFAULT);
        DataColumn.Value renamed = val.withName("renamed");
        assertThat(renamed.getName()).isEqualTo("renamed");
        assertThat(renamed.getStyle()).isEqualTo(val.getStyle());
    }

    @Test
    void dataColumnWithDefaults() {
        DataColumn.Value field = new DataColumn.Value("field", "fieldStyle", "", 10, OptBoolean.DEFAULT, 0, 0, OptBoolean.DEFAULT);
        DataGrid.Value grid = new DataGrid.Value("gridStyle", "gridHeaderStyle", 20, OptBoolean.TRUE, 5, 50, OptBoolean.TRUE);
        DataColumn.Value result = field.withDefaults(grid);

        assertThat(result.getStyle()).isEqualTo("fieldStyle");
        assertThat(result.getWidth()).isEqualTo(10);
        // width=10 → isAutoSize() is false (requires width < 0)
        assertThat(result.isAutoSize()).isFalse();
        assertThat(result.isMerge()).isTrue();
    }
}
