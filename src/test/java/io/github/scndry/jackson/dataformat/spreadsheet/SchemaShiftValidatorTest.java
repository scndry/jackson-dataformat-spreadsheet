package io.github.scndry.jackson.dataformat.spreadsheet;

import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThatCode;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Schema-build-time validation for the {@code shift} attribute on
 * {@code @DataColumn} and {@code @DataColumnGroup}. Invariant:
 * {@code shift >= 0}.
 */
class SchemaShiftValidatorTest {

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NoShift {
        @DataColumn("name") String name;
        @DataColumn("qty") int qty;
    }

    @Test
    void zeroShift_withHeader_doesNotThrow() {
        assertThatCode(() -> new SpreadsheetMapper().sheetSchemaFor(NoShift.class))
                .doesNotThrowAnyException();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class WithShift {
        @DataColumn("name") String name;
        @DataColumn(value = "total", shift = 1) double total;
    }

    @Test
    void positiveShift_withoutHeader_doesNotThrow() {
        final SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .useHeader(false)
                .build();
        assertThatCode(() -> mapper.sheetSchemaFor(WithShift.class))
                .doesNotThrowAnyException();
    }

    @Test
    void positiveShift_withHeader_doesNotThrow() {
        assertThatCode(() -> new SpreadsheetMapper().sheetSchemaFor(WithShift.class))
                .doesNotThrowAnyException();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NegativeShift {
        @DataColumn("name") String name;
        @DataColumn(value = "total", shift = -1) double total;
    }

    @Test
    void negativeShift_throws() {
        final SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .useHeader(false)
                .build();
        assertThatThrownBy(() -> mapper.sheetSchemaFor(NegativeShift.class))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("shift")
                .hasMessageContaining("must be >= 0");
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Inner {
        @DataColumn("product") String product;
        @DataColumn("qty") int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupWithShift {
        @DataColumn("orderId") int orderId;
        @DataColumnGroup(value = "Items", shift = 1) List<Inner> items;
    }

    @Test
    void groupShift_withoutHeader_doesNotThrow() {
        final SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .useHeader(false)
                .build();
        assertThatCode(() -> mapper.sheetSchemaFor(GroupWithShift.class))
                .doesNotThrowAnyException();
    }

    @Test
    void groupShift_withHeader_doesNotThrow() {
        assertThatCode(() -> new SpreadsheetMapper().sheetSchemaFor(GroupWithShift.class))
                .doesNotThrowAnyException();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupWithNegativeShift {
        @DataColumn("orderId") int orderId;
        @DataColumnGroup(value = "Items", shift = -1) List<Inner> items;
    }

    @Test
    void groupShift_negative_throws() {
        final SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .useHeader(false)
                .build();
        assertThatThrownBy(() -> mapper.sheetSchemaFor(GroupWithNegativeShift.class))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("shift")
                .hasMessageContaining("must be >= 0");
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class ExcessiveShift {
        @DataColumn("name") String name;
        @DataColumn(value = "x", shift = 20000) int x;
    }

    @Test
    void excessiveShift_throws() {
        final SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .useHeader(false)
                .build();
        assertThatThrownBy(() -> mapper.sheetSchemaFor(ExcessiveShift.class))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("'x'")
                .hasMessageContaining("must be <= 16384");
    }
}
