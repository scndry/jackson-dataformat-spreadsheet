package io.github.scndry.jackson.dataformat.spreadsheet;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.Data;
import lombok.NoArgsConstructor;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

/**
 * Nested-list bench model — 10 columns total, same type distribution as
 * {@link BenchRow}, split into 6 outer + 4 inner. {@code id} is the row
 * anchor; {@code totalAmount} and {@code createdAt} sit after the items
 * list to exercise the outer-after-list back-write path on write and
 * the record-tree close on read.
 */
@Data @NoArgsConstructor @DataGrid
public class NestedRow {

    @DataColumn(merge = OptBoolean.TRUE, anchor = OptBoolean.TRUE)
    private Long id;
    @DataColumn(merge = OptBoolean.TRUE)
    private String name;
    @DataColumn(merge = OptBoolean.TRUE)
    private String category;
    @DataColumn(merge = OptBoolean.TRUE)
    private String status;

    @DataColumnGroup("Items")
    private List<NestedLineItem> items;

    @DataColumn(merge = OptBoolean.TRUE)
    private BigDecimal totalAmount;
    @DataColumn(merge = OptBoolean.TRUE)
    private LocalDateTime createdAt;

    private static final String[] CATEGORIES = {
            "Electronics", "Clothing", "Food", "Books", "Sports",
            "Home", "Garden", "Toys", "Health", "Automotive"
    };

    private static final String[] STATUSES = {
            "Active", "Inactive", "Pending", "Archived"
    };

    public static NestedRow create(final int outerIdx, final int itemsPerOuter) {
        NestedRow r = new NestedRow();
        r.setId((long) outerIdx);
        r.setName("Product-" + outerIdx);
        r.setCategory(CATEGORIES[outerIdx % CATEGORIES.length]);
        r.setStatus(STATUSES[outerIdx % STATUSES.length]);
        r.setTotalAmount(BigDecimal.valueOf(100.50 + outerIdx * 0.01));
        r.setCreatedAt(LocalDateTime.of(2024, 1, 1, 0, 0).plusMinutes(outerIdx));
        List<NestedLineItem> items = new ArrayList<>(itemsPerOuter);
        final int base = outerIdx * itemsPerOuter;
        for (int j = 0; j < itemsPerOuter; j++) {
            items.add(NestedLineItem.create(base + j));
        }
        r.setItems(items);
        return r;
    }

    public static List<NestedRow> sample(final int outerCount, final int itemsPerOuter) {
        List<NestedRow> list = new ArrayList<>(outerCount);
        for (int i = 0; i < outerCount; i++) {
            list.add(create(i, itemsPerOuter));
        }
        return list;
    }
}
