package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

@Data @NoArgsConstructor @DataGrid
public class BenchRow {
    private Long id;
    private String name;
    private String category;
    private String status;
    private int quantity;
    private double price;
    private BigDecimal amount;
    private LocalDate dueDate;
    private String description;
    private LocalDateTime createdAt;

    public static final String[] HEADERS = {
            "id", "name", "category", "status", "quantity",
            "price", "amount", "dueDate", "description", "createdAt"
    };

    private static final String[] CATEGORIES = {
            "Electronics", "Clothing", "Food", "Books", "Sports",
            "Home", "Garden", "Toys", "Health", "Automotive"
    };

    private static final String[] STATUSES = {
            "Active", "Inactive", "Pending", "Archived"
    };

    public static BenchRow create(final int i) {
        BenchRow r = new BenchRow();
        r.setId((long) i);
        r.setName("Product-" + i);
        r.setCategory(CATEGORIES[i % CATEGORIES.length]);
        r.setStatus(STATUSES[i % STATUSES.length]);
        r.setQuantity(i % 1000);
        r.setPrice(9.99 + (i % 1000) * 0.5);
        r.setAmount(BigDecimal.valueOf(100.50 + i * 0.01));
        r.setDueDate(LocalDate.of(2024, 1, 1).plusDays(i % 365));
        r.setDescription("description-of-product-" + i);
        r.setCreatedAt(LocalDateTime.of(2024, 1, 1, 0, 0).plusMinutes(i));
        return r;
    }
}
