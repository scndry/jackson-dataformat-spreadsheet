package io.github.scndry.jackson.dataformat.spreadsheet.internal;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

/**
 * Internal profiling benchmark — single outer record with a large
 * inner list. Exercises the write-side back-write buffer (outer
 * field after the list, back-written into the row already emitted
 * for {@code items[0]}) and the read-side record-tree assembly
 * (single outer absorbs N inner rows). Used for optimization work;
 * not documented in BENCHMARK.md.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(value = 1, jvmArgs = {"-Xmx4g"})
public class LargeInnerListBenchmark {

    @Param({"1000", "10000", "100000"})
    int innerCount;

    Order order;
    File readFixture;
    File writeTarget;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid(mergeColumn = OptBoolean.TRUE)
    public static class Order {
        @DataColumn(value = "id", anchor = true) int id;
        @DataColumn("customer") String customer;
        @DataColumnGroup("items") List<Item> items;
        @DataColumn("total") double total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    public static class Item {
        @DataColumn("product") String product;
        @DataColumn("qty") int qty;
        @DataColumn("amount") double amount;
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        final List<Item> items = new ArrayList<>(innerCount);
        for (int i = 0; i < innerCount; i++) {
            items.add(new Item("sku-" + i, i, i * 1.5));
        }
        order = new Order(1, "Alice", items, 99_999.99);
        readFixture = File.createTempFile("bench-large-inner-list-read-", ".xlsx");
        readFixture.deleteOnExit();
        new SpreadsheetMapper().writeValue(readFixture, order, Order.class);
    }

    @TearDown(Level.Trial)
    public void tearDownTrial() {
        readFixture.delete();
    }

    @Setup(Level.Invocation)
    public void setUpFile() throws IOException {
        writeTarget = File.createTempFile("bench-large-inner-list-write-", ".xlsx");
        writeTarget.deleteOnExit();
    }

    @TearDown(Level.Invocation)
    public void tearDown() {
        writeTarget.delete();
    }

    @Benchmark
    public void write(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(writeTarget, order, Order.class);
        bh.consume(writeTarget);
    }

    @Benchmark
    public void read(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        Order result = mapper.readValue(readFixture, Order.class);
        bh.consume(result);
    }
}
