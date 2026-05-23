package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonView;
import com.fasterxml.jackson.annotation.OptBoolean;
import com.fasterxml.jackson.databind.MapperFeature;
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

/**
 * Round-trip tests for nested-list read/write — single record,
 * 1/2/3-depth nesting, sibling lists, empty list, {@code @JsonView}
 * filtering, and SSML/POI cross-mode parity. Common-case shapes
 * live here; partial-cell edge cases in
 * {@link NestedListEdgeCaseTest}, blank-row semantics in
 * {@link NestedBlankRowFeatureTest}.
 */
class NestedListReadTest {

    @TempDir Path tempDir;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Order {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE, anchor = true) int id;
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
        @DataColumn(value = "OID", merge = OptBoolean.TRUE, anchor = true) int id;
        @DataColumnGroup("Mid") List<MidLevel> items;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidLevel {
        @DataColumn(value = "MID", merge = OptBoolean.TRUE, anchor = true) String midId;
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
        @DataColumn(value = "TID", merge = OptBoolean.TRUE, anchor = true) int id;
        @DataColumnGroup("Mid") List<MidNode> mids;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidNode {
        @DataColumn(value = "MID", merge = OptBoolean.TRUE, anchor = true) String mid;
        @DataColumnGroup("Sub") List<SubNode> subs;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class SubNode {
        @DataColumn(value = "SID", merge = OptBoolean.TRUE, anchor = true) String sid;
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
        @DataColumn(value = "ID", merge = OptBoolean.TRUE, anchor = true) int id;
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
        @DataColumn(value = "GID", merge = OptBoolean.TRUE, anchor = true) String gid;
        @DataColumnGroup("Detail") List<Inner> details;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Receipt {
        @DataColumn("Note") String note;
        @DataColumn("Total") int total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Invoice {
        @DataColumn(value = "INV", merge = OptBoolean.TRUE, anchor = true) int id;
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
        @DataColumn(value = "MID", anchor = true, merge = OptBoolean.TRUE) String mid;
        @DataColumnGroup("As") List<Inner> as;
        @DataColumnGroup("Bs") List<Inner> bs;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NestedSiblingTop {
        @DataColumn(value = "ID", anchor = true, merge = OptBoolean.TRUE) int id;
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

    // POI User Model parity — RecordTreeBuffer sits at the parser level so the
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

    /**
     * Three-depth fixture round-tripped across SSML and POI modes —
     * SSML write then POI read, POI write then SSML read — confirms
     * the deepest nesting layout is reader-agnostic.
     */
    @Test
    void roundTrip_threeDepth_crossPoiSsml() throws Exception {
        SpreadsheetMapper ssml = new SpreadsheetMapper();
        SpreadsheetMapper poi = _poiMapper();

        TopLevel expected = new TopLevel(1, Arrays.asList(
                new MidNode("M1", Arrays.asList(
                        new SubNode("S1A", Arrays.asList(new Inner(1, 2), new Inner(3, 4))),
                        new SubNode("S1B", Arrays.asList(new Inner(5, 6))))),
                new MidNode("M2", Arrays.asList(
                        new SubNode("S2A", Arrays.asList(new Inner(7, 8)))))));

        File ssmlFile = tempDir.resolve("three-depth-ssml-write.xlsx").toFile();
        ssml.writeValue(ssmlFile, expected);
        List<TopLevel> readViaPoi = poi.readValues(ssmlFile, TopLevel.class);
        assertThat(readViaPoi).hasSize(1);
        assertThat(readViaPoi.get(0)).isEqualTo(expected);

        File poiFile = tempDir.resolve("three-depth-poi-write.xlsx").toFile();
        poi.writeValue(poiFile, expected);
        List<TopLevel> readViaSsml = ssml.readValues(poiFile, TopLevel.class);
        assertThat(readViaSsml).hasSize(1);
        assertThat(readViaSsml.get(0)).isEqualTo(expected);
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ComplexMid {
        @DataColumn(value = "CMID", anchor = true, merge = OptBoolean.TRUE) String mid;
        @DataColumnGroup("Subs") List<SubNode> subs;
        @DataColumnGroup("Notes") List<Inner> notes;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class ComplexTop {
        @DataColumn(value = "CTID", anchor = true, merge = OptBoolean.TRUE) int id;
        @DataColumnGroup("Mids") List<ComplexMid> mids;
    }

    /**
     * Three-depth path (Top -> Mid -> Sub -> Leaf) combined with a
     * sibling list at the Mid level (Subs alongside Notes). Exercises
     * the record-tree close path when both a deep descendant and a
     * sibling close at the same scope.
     */
    @Test
    void roundTrip_threeDepthWithSibling() throws Exception {
        File file = tempDir.resolve("three-depth-sibling.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        // Sibling lists must share row span (write-side invariant):
        // M1.subs leaves total 3 -> M1.notes also 3 items.
        // M2.subs leaves total 1 -> M2.notes also 1 item.
        ComplexTop expected = new ComplexTop(1, Arrays.asList(
                new ComplexMid("M1",
                        Arrays.asList(
                                new SubNode("S1A", Arrays.asList(new Inner(1, 2), new Inner(3, 4))),
                                new SubNode("S1B", Arrays.asList(new Inner(5, 6)))),
                        Arrays.asList(new Inner(10, 20), new Inner(30, 40), new Inner(50, 60))),
                new ComplexMid("M2",
                        Arrays.asList(
                                new SubNode("S2A", Arrays.asList(new Inner(7, 8)))),
                        Arrays.asList(new Inner(70, 80)))));

        mapper.writeValue(file, expected);

        List<ComplexTop> read = mapper.readValues(file, ComplexTop.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0)).isEqualTo(expected);
    }

    interface Views {
        interface Summary {}
        interface Detail extends Summary {}
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ViewItem {
        @DataColumn("sku") @JsonView(Views.Summary.class) String sku;
        @DataColumn("qty") @JsonView(Views.Detail.class) Integer qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class ViewOrder {
        @DataColumn(value = "id", anchor = true, merge = OptBoolean.TRUE)
        @JsonView(Views.Summary.class) Integer id;
        @DataColumnGroup("Items")
        @JsonView(Views.Summary.class) List<ViewItem> items;
        @DataColumn("total")
        @JsonView(Views.Detail.class) Integer total;
    }

    /**
     * {@code @JsonView} + nested round-trip. Summary view writes
     * anchor + items.sku only; Detail-only fields (items.qty, total)
     * are absent from the sheet, so the subsequent default read sees
     * them as {@code null}.
     */
    @Test
    void roundTrip_jsonView_nested_summary() throws Exception {
        SpreadsheetMapper viewMapper = SpreadsheetMapper.builder()
                .configure(MapperFeature.DEFAULT_VIEW_INCLUSION, false)
                .build();
        File file = tempDir.resolve("nested-view-summary.xlsx").toFile();

        ViewOrder data = new ViewOrder(1,
                Arrays.asList(new ViewItem("A", 10), new ViewItem("B", 20)), 100);
        viewMapper.sheetWriterForWithView(ViewOrder.class, Views.Summary.class)
                .writeValue(file, Arrays.asList(data));

        List<ViewOrder> read = viewMapper.readValues(file, ViewOrder.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0).getId()).isEqualTo(1);
        assertThat(read.get(0).getItems()).hasSize(2);
        assertThat(read.get(0).getItems()).extracting(ViewItem::getSku)
                .containsExactly("A", "B");
        // Detail-only fields not written under Summary view
        assertThat(read.get(0).getItems()).extracting(ViewItem::getQty)
                .containsExactly(null, null);
        assertThat(read.get(0).getTotal()).isNull();
    }
}
