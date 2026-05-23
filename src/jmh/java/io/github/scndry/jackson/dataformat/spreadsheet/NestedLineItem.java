package io.github.scndry.jackson.dataformat.spreadsheet;

import java.time.LocalDate;

import lombok.Data;
import lombok.NoArgsConstructor;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;

@Data @NoArgsConstructor
public class NestedLineItem {

    @DataColumn private int quantity;
    @DataColumn private double price;
    @DataColumn private LocalDate dueDate;
    @DataColumn private String description;

    public static NestedLineItem create(final int i) {
        NestedLineItem item = new NestedLineItem();
        item.setQuantity(i % 1000);
        item.setPrice(9.99 + (i % 1000) * 0.5);
        item.setDueDate(LocalDate.of(2024, 1, 1).plusDays(i % 365));
        item.setDescription("line-" + i);
        return item;
    }
}
