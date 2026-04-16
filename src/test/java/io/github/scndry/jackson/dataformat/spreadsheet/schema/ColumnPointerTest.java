package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.stream.Collectors;

import static org.assertj.core.api.Assertions.*;

class ColumnPointerTest {

    @Test void emptyIsEmpty()     { assertThat(ColumnPointer.empty().isEmpty()).isTrue(); }
    @Test void emptyDepthIsOne()  { assertThat(ColumnPointer.empty().depth()).isEqualTo(1); }
    @Test void arrayNotEmpty()    { assertThat(ColumnPointer.array().isEmpty()).isFalse(); }
    @Test void parentIsParent()   { assertThat(SegmentPointer.PARENT.isParent()).isTrue(); }

    @Test void resolveFromEmpty() {
        ColumnPointer p = ColumnPointer.empty().resolve("a");
        assertThat(p.toString()).isEqualTo("a");
        assertThat(p.depth()).isEqualTo(1);
    }

    @Test void resolveChain() {
        ColumnPointer p = ColumnPointer.empty().resolve("a").resolve("b").resolve("c");
        assertThat(p.toString()).isEqualTo("a/b/c");
        assertThat(p.depth()).isEqualTo(3);
    }

    @Test void resolvePointer() {
        ColumnPointer a = of("a");
        ColumnPointer bc = of("b", "c");
        assertThat(a.resolve(bc).toString()).isEqualTo("a/b/c");
    }

    @Test void resolveArray() {
        assertThat(of("a").resolveArray().toString()).isEqualTo("a/[]");
    }

    @Test void relativizeChild() {
        assertThat(of("a", "b").relativize(of("a", "b", "c", "d")).toString()).isEqualTo("c/d");
    }

    @Test void relativizeSibling() {
        assertThat(of("a", "b", "c").relativize(of("a", "d")).toString()).isEqualTo("../../d");
    }

    @Test void relativizeSelf() {
        assertThat(of("a").relativize(of("a")).toString()).isEmpty();
    }

    @Test void relativizeFromEmpty() {
        assertThat(ColumnPointer.empty().relativize(of("a", "b")).toString()).isEqualTo("a/b");
    }

    @Test void getParentNested() { assertThat(of("a", "b").getParent().toString()).isEqualTo("a"); }
    @Test void getParentSingle() { assertThat(of("a").getParent().isEmpty()).isTrue(); }
    @Test void getParentEmpty()  { assertThat(ColumnPointer.empty().getParent().isEmpty()).isTrue(); }

    @Test void head() { assertThat(of("a", "b", "c").head().toString()).isEqualTo("a"); }
    @Test void name() { assertThat(of("a", "b", "c").name()).isEqualTo("c"); }

    @Test void startsWithPrefix()  { assertThat(of("a", "b", "c").startsWith(of("a", "b"))).isTrue(); }
    @Test void startsWithLonger()  { assertThat(of("a", "b").startsWith(of("a", "b", "c"))).isFalse(); }
    @Test void startsWithNoMatch() { assertThat(of("a").startsWith(of("x"))).isFalse(); }
    @Test void startsWithEmpty()   { assertThat(of("a").startsWith(ColumnPointer.empty())).isTrue(); }

    @Test void containsYes() { assertThat(of("a", "b", "c").contains(of("b"))).isTrue(); }
    @Test void containsNo()  { assertThat(of("a", "b").contains(of("x"))).isFalse(); }

    @Test void stream() {
        List<String> names = of("a", "b", "c").stream()
                .map(ColumnPointer::toString).collect(Collectors.toList());
        assertThat(names).containsExactly("a", "b", "c");
    }

    @Test void emptyStream() {
        assertThat(ColumnPointer.empty().stream().count()).isZero();
    }

    private static ColumnPointer of(String... segments) {
        ColumnPointer r = ColumnPointer.empty();
        for (String s : segments) r = r.resolve(s);
        return r;
    }
}
