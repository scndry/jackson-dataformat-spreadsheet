package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class FormulaTest {

    @Test
    void of_storesTextAndKind() {
        Formula f = Formula.of("$D$1");
        assertThat(f.kind()).isEqualTo(Formula.Kind.OF);
        assertThat(f.value()).isEqualTo("$D$1");
    }

    @Test
    void of_acceptsExpressionString() {
        Formula f = Formula.of("AVERAGE($B$2:$B$100) * 0.9");
        assertThat(f.kind()).isEqualTo(Formula.Kind.OF);
        assertThat(f.value()).isEqualTo("AVERAGE($B$2:$B$100) * 0.9");
    }

    @Test
    void of_rejectsNull() {
        assertThatThrownBy(() -> Formula.of(null))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void of_rejectsEmpty() {
        assertThatThrownBy(() -> Formula.of(""))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void column_storesNameAndKind() {
        Formula f = Formula.column("minPrice");
        assertThat(f.kind()).isEqualTo(Formula.Kind.COLUMN);
        assertThat(f.value()).isEqualTo("minPrice");
    }

    @Test
    void column_rejectsNull() {
        assertThatThrownBy(() -> Formula.column(null))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void column_rejectsEmpty() {
        assertThatThrownBy(() -> Formula.column(""))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("must not be empty");
    }
}
