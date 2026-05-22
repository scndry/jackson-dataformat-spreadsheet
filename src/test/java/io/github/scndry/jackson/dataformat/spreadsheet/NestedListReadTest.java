package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

class NestedListReadTest {

    @TempDir Path tempDir;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Order {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) int id;
        @DataColumn(value = "Customer", merge = OptBoolean.TRUE) String customer;
        @DataColumnGroup("Items") List<LineItem> items;
        @DataColumn(value = "Total", merge = OptBoolean.TRUE) int total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class LineItem {
        @DataColumn("SKU") String sku;
        @DataColumn("Qty") int qty;
        @DataColumn("Amount") int amount;
    }

    @Test
    void roundTrip_singleOrder() throws Exception {
        File file = tempDir.resolve("single-order.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        Order expected = new Order(1, "Alice",
                Arrays.asList(new LineItem("A-01", 3, 30), new LineItem("B-02", 5, 65)),
                95);

        mapper.writeValue(file, expected);

        List<Order> read = mapper.readValues(file, Order.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    @Test
    void roundTrip_multipleOrders() throws Exception {
        File file = tempDir.resolve("multi-order.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        List<Order> expected = Arrays.asList(
                new Order(1, "Alice", Arrays.asList(
                        new LineItem("A-01", 3, 30),
                        new LineItem("B-02", 5, 65)),
                        95),
                new Order(2, "Bob", Arrays.asList(
                        new LineItem("C-03", 1, 10)),
                        10));

        mapper.writeValue(file, expected, Order.class);

        List<Order> read = mapper.readValues(file, Order.class);
        assertThat(read).isEqualTo(expected);
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class DeepOuter {
        @DataColumn(value = "OID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) int id;
        @DataColumnGroup("Mid") List<MidLevel> items;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidLevel {
        @DataColumn(value = "MID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) String midId;
        @DataColumnGroup("Detail") List<Inner> details;
        @DataColumn(value = "MNAME", merge = OptBoolean.TRUE) String midName;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Inner {
        @DataColumn("X") int x;
        @DataColumn("Y") int y;
    }

      @Test
    void roundTrip_twoDepth() throws Exception {
        File file = tempDir.resolve("two-depth.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        DeepOuter expected = new DeepOuter(1, Arrays.asList(
                new MidLevel("X", Arrays.asList(new Inner(1, 2), new Inner(3, 4)), "Name-X"),
                new MidLevel("Y", Arrays.asList(new Inner(5, 6)), "Name-Y")));

        mapper.writeValue(file, expected);

        List<DeepOuter> read = mapper.readValues(file, DeepOuter.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class TopLevel {
        @DataColumn(value = "TID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) int id;
        @DataColumnGroup("Mid") List<MidNode> mids;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidNode {
        @DataColumn(value = "MID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) String mid;
        @DataColumnGroup("Sub") List<SubNode> subs;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class SubNode {
        @DataColumn(value = "SID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) String sid;
        @DataColumnGroup("Leaf") List<Inner> leaves;
    }

    @Test
    void roundTrip_threeDepth() throws Exception {
        File file = tempDir.resolve("three-depth.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        TopLevel expected = new TopLevel(1, Arrays.asList(
                new MidNode("M1", Arrays.asList(
                        new SubNode("S1A", Arrays.asList(new Inner(1, 2), new Inner(3, 4))),
                        new SubNode("S1B", Arrays.asList(new Inner(5, 6))))),
                new MidNode("M2", Arrays.asList(
                        new SubNode("S2A", Arrays.asList(new Inner(7, 8)))))));

        mapper.writeValue(file, expected);

        List<TopLevel> read = mapper.readValues(file, TopLevel.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class SiblingItem {
        @DataColumn("SKU") String sku;
        @DataColumn("Qty") int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class SiblingPayment {
        @DataColumn("Method") String method;
        @DataColumn("Amount") int amount;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class SiblingOrder {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) int id;
        @DataColumnGroup("Items") List<SiblingItem> items;
        @DataColumnGroup("Payments") List<SiblingPayment> payments;
    }

    @Test
    void roundTrip_siblingLists() throws Exception {
        File file = tempDir.resolve("sibling.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        SiblingOrder expected = new SiblingOrder(1,
                Arrays.asList(new SiblingItem("A", 10), new SiblingItem("B", 20)),
                Arrays.asList(new SiblingPayment("card", 100), new SiblingPayment("cash", 200)));

        mapper.writeValue(file, expected);

        List<SiblingOrder> read = mapper.readValues(file, SiblingOrder.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class GroupItem {
        @DataColumn(value = "GID", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) String gid;
        @DataColumnGroup("Detail") List<Inner> details;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Receipt {
        @DataColumn("Note") String note;
        @DataColumn("Total") int total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Invoice {
        @DataColumn(value = "INV", merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE) int id;
        @DataColumnGroup("Groups") List<GroupItem> groups;
        @DataColumnGroup("Receipts") List<Receipt> receipts;
    }

    @Test
    void roundTrip_siblingWithNested() throws Exception {
        File file = tempDir.resolve("sibling-nested.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        Invoice expected = new Invoice(1,
                Arrays.asList(
                        new GroupItem("G1", Arrays.asList(new Inner(1, 2), new Inner(3, 4))),
                        new GroupItem("G2", Arrays.asList(new Inner(5, 6)))),
                Arrays.asList(
                        new Receipt("first", 10),
                        new Receipt("second", 20),
                        new Receipt("third", 30)));

        mapper.writeValue(file, expected);

        List<Invoice> read = mapper.readValues(file, Invoice.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    // Empty list — write produces zero inner rows; Jackson convention
    // makes the missing array field deserialize as null (use
    // @JsonSetter(nulls = AS_EMPTY) on the field for empty-list semantics).
    @Test
    void roundTrip_emptyList() throws Exception {
        File file = tempDir.resolve("empty-list.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        Order written = new Order(1, "Alice", Collections.emptyList(), 0);
        mapper.writeValue(file, written);

        List<Order> read = mapper.readValues(file, Order.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(new Order(1, "Alice", null, 0));
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class NestedSiblingMid {
        @DataColumn(value = "MID", anchor = OptBoolean.TRUE, merge = OptBoolean.TRUE) String mid;
        @DataColumnGroup("As") List<Inner> as;
        @DataColumnGroup("Bs") List<Inner> bs;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NestedSiblingTop {
        @DataColumn(value = "ID", anchor = OptBoolean.TRUE, merge = OptBoolean.TRUE) int id;
        @DataColumnGroup("Mids") List<NestedSiblingMid> mids;
    }

    @Test
    void roundTrip_nestedSiblingLists() throws Exception {
        File file = tempDir.resolve("nested-sibling.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        NestedSiblingTop expected = new NestedSiblingTop(1, Arrays.asList(
                new NestedSiblingMid("M1",
                        Arrays.asList(new Inner(1, 2), new Inner(3, 4)),
                        Arrays.asList(new Inner(11, 12), new Inner(13, 14))),
                new NestedSiblingMid("M2",
                        Arrays.asList(new Inner(5, 6)),
                        Arrays.asList(new Inner(15, 16)))));

        mapper.writeValue(file, expected);

        List<NestedSiblingTop> read = mapper.readValues(file, NestedSiblingTop.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    // POI User Model parity — NestedReadAlg sits at the parser level so the
    // read path (SSML vs POI object model) should be transparent. Spot-check
    // one case per shape (single-record, 2-depth + outer-after-list, sibling).

    private SpreadsheetMapper _poiMapper() {
        return new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
    }

    @Test
    void roundTrip_singleOrder_poi() throws Exception {
        File file = tempDir.resolve("single-order-poi.xlsx").toFile();
        SpreadsheetMapper mapper = _poiMapper();

        Order expected = new Order(1, "Alice",
                Arrays.asList(new LineItem("A-01", 3, 30), new LineItem("B-02", 5, 65)),
                95);

        mapper.writeValue(file, expected);

        List<Order> read = mapper.readValues(file, Order.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    @Test
    void roundTrip_twoDepth_poi() throws Exception {
        File file = tempDir.resolve("two-depth-poi.xlsx").toFile();
        SpreadsheetMapper mapper = _poiMapper();

        DeepOuter expected = new DeepOuter(1, Arrays.asList(
                new MidLevel("X", Arrays.asList(new Inner(1, 2), new Inner(3, 4)), "Name-X"),
                new MidLevel("Y", Arrays.asList(new Inner(5, 6)), "Name-Y")));

        mapper.writeValue(file, expected);

        List<DeepOuter> read = mapper.readValues(file, DeepOuter.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    @Test
    void roundTrip_siblingLists_poi() throws Exception {
        File file = tempDir.resolve("sibling-poi.xlsx").toFile();
        SpreadsheetMapper mapper = _poiMapper();

        SiblingOrder expected = new SiblingOrder(1,
                Arrays.asList(new SiblingItem("A", 10), new SiblingItem("B", 20)),
                Arrays.asList(new SiblingPayment("card", 100), new SiblingPayment("cash", 200)));

        mapper.writeValue(file, expected);

        List<SiblingOrder> read = mapper.readValues(file, SiblingOrder.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }
}
