package io.github.scndry.jackson.dataformat.spreadsheet;

import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.NestedAnchorValidator;

import static org.assertj.core.api.Assertions.assertThatCode;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Read-time anchor invariant checks — exactly one
 * {@code @DataColumn(anchor = TRUE)} at each nested-list-bearing record
 * level, none elsewhere. Driven from {@code SheetParser.setSchema}
 * when the schema has any anchor; the write path is untouched.
 */
class NestedAnchorValidatorTest {

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Inner {
        @DataColumn("x") int x;
        @DataColumn("y") int y;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Valid {
        @DataColumn(value = "id", anchor = OptBoolean.TRUE) int id;
        List<Inner> items;
    }

    @Test
    void valid_anchorAtListBearingScope_doesNotThrow() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Valid.class);
        assertThatCode(() -> NestedAnchorValidator.validate(schema)).doesNotThrowAnyException();
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidNoAnchor {
        @DataColumn("midId") String midId;
        List<Inner> details;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class MissingAtNested {
        @DataColumn(value = "id", anchor = OptBoolean.TRUE) int id;
        List<MidNoAnchor> items;
    }

    @Test
    void missing_anchorAtNestedScope_throws() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(MissingAtNested.class);
        assertThatThrownBy(() -> NestedAnchorValidator.validate(schema))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Missing anchor")
                .hasMessageContaining("items/[]");
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidWithAnchor {
        @DataColumn(value = "midId", anchor = OptBoolean.TRUE) String midId;
        List<Inner> details;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class MissingAtRoot {
        @DataColumn("id") int id;
        List<MidWithAnchor> items;
    }

    @Test
    void missing_anchorAtRootScope_throws() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(MissingAtRoot.class);
        assertThatThrownBy(() -> NestedAnchorValidator.validate(schema))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Missing anchor")
                .hasMessageContaining("<root>");
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class MissingAll {
        @DataColumn("id") int id;
        List<Inner> items;
    }

    @Test
    void missing_noAnchorAnywhereWithNestedList_throws() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(MissingAll.class);
        assertThatThrownBy(() -> NestedAnchorValidator.validate(schema))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Missing anchor")
                .hasMessageContaining("<root>");
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class InnerWithAnchor {
        @DataColumn(value = "x", anchor = OptBoolean.TRUE) int x;
        @DataColumn("y") int y;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class ExtraAtInnermost {
        @DataColumn(value = "id", anchor = OptBoolean.TRUE) int id;
        List<InnerWithAnchor> items;
    }

    @Test
    void extra_anchorAtInnermostScope_throws() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(ExtraAtInnermost.class);
        assertThatThrownBy(() -> NestedAnchorValidator.validate(schema))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Extra anchor")
                .hasMessageContaining("items/[]");
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class ExtraOnFlat {
        @DataColumn(value = "id", anchor = OptBoolean.TRUE) int id;
        @DataColumn("name") String name;
    }

    @Test
    void extra_anchorOnFlatSchema_throws() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(ExtraOnFlat.class);
        assertThatThrownBy(() -> NestedAnchorValidator.validate(schema))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Extra anchor")
                .hasMessageContaining("<root>");
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class DuplicateAtRoot {
        @DataColumn(value = "id", anchor = OptBoolean.TRUE) int id;
        @DataColumn(value = "name", anchor = OptBoolean.TRUE) String name;
        List<Inner> items;
    }

    @Test
    void duplicate_anchorAtSameScope_throws() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(DuplicateAtRoot.class);
        assertThatThrownBy(() -> NestedAnchorValidator.validate(schema))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Multiple anchors")
                .hasMessageContaining("id")
                .hasMessageContaining("name");
    }

    private static SpreadsheetSchema _schemaFor(final Class<?> type) throws Exception {
        return new SpreadsheetMapper().sheetSchemaFor(type);
    }
}
