package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Direct tests for {@link InMemorySharedStringsStore} — covers branches
 * that integration tests don't reach: null input, dedup hit, resize at
 * load factor, value-array expansion, char slab expansion.
 */
class InMemorySharedStringsStoreTest {

    @Test
    void getOrAdd_null_treatedAsEmpty() {
        InMemorySharedStringsStore store = new InMemorySharedStringsStore();
        int idx = store.getOrAdd(null);
        assertThat(store.get(idx)).isEqualTo("");
        assertThat(store.size()).isEqualTo(1);
    }

    @Test
    void getOrAdd_duplicate_returnsSameIndex() {
        InMemorySharedStringsStore store = new InMemorySharedStringsStore();
        int first = store.getOrAdd("hello");
        int second = store.getOrAdd("hello");
        assertThat(first).isEqualTo(second);
        assertThat(store.size()).isEqualTo(1);
    }

    @Test
    void getOrAdd_resizeAndExpansion_preservesAllValues() {
        // initial capacity 16: triggers index-table resize, value-array growth,
        // and char-slab growth as 200 unique strings are added.
        InMemorySharedStringsStore store = new InMemorySharedStringsStore(16);
        for (int i = 0; i < 200; i++) {
            store.getOrAdd("value-" + i);
        }
        assertThat(store.size()).isEqualTo(200);
        assertThat(store.get(0)).isEqualTo("value-0");
        assertThat(store.get(199)).isEqualTo("value-199");
        // Re-add an existing entry — must hit the dedup path, not append
        int existingIdx = store.getOrAdd("value-100");
        assertThat(existingIdx).isEqualTo(100);
        assertThat(store.size()).isEqualTo(200);
    }

    @Test
    void getOrAdd_largeStringTriggersCharSlabExpansion() {
        InMemorySharedStringsStore store = new InMemorySharedStringsStore();
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < 5000; i++) sb.append('x');  // exceeds INITIAL_CHAR_CAPACITY 4096
        int idx = store.getOrAdd(sb.toString());
        assertThat(store.length(idx)).isEqualTo(5000);
        assertThat(store.get(idx)).isEqualTo(sb.toString());
    }

    @Test
    void charAt_returnsCharAtOffset() {
        InMemorySharedStringsStore store = new InMemorySharedStringsStore();
        int idx = store.getOrAdd("hello");
        assertThat(store.charAt(idx, 0)).isEqualTo('h');
        assertThat(store.charAt(idx, 4)).isEqualTo('o');
    }

    @Test
    void close_isNoOp() throws Exception {
        InMemorySharedStringsStore store = new InMemorySharedStringsStore();
        store.getOrAdd("x");
        store.close();
        assertThat(store.size()).isEqualTo(1);
        assertThat(store.get(0)).isEqualTo("x");
    }
}
