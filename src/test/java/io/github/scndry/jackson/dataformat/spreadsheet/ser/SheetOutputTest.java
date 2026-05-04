package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SheetOutputTest {

    private static final File DUMMY_FILE = new File("dummy.xlsx");
    private static final Path DUMMY_PATH = DUMMY_FILE.toPath();
    private static final ByteArrayOutputStream DUMMY_STREAM = new ByteArrayOutputStream();

    @Test
    void allowsNullName() {
        assertThat(SheetOutput.target(DUMMY_FILE, null).getName()).isNull();
        assertThat(SheetOutput.target(DUMMY_PATH, null).getName()).isNull();
        assertThat(SheetOutput.target(DUMMY_STREAM, null).getName()).isNull();
    }

    @Test
    void acceptsValidName() {
        assertThat(SheetOutput.target(DUMMY_FILE, "A").getName()).isEqualTo("A");
        assertThat(SheetOutput.target(DUMMY_FILE, "1234567890123456789012345678901").getName())
                .hasSize(31);
    }

    @Test
    void rejectsEmptyName() {
        assertThatThrownBy(() -> SheetOutput.target(DUMMY_FILE, ""))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void rejectsTooLongName() {
        // 32 chars
        assertThatThrownBy(() -> SheetOutput.target(DUMMY_FILE, "12345678901234567890123456789012"))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void rejectsForbiddenCharacters() {
        for (final char c : new char[]{'/', '\\', '?', '*', '[', ']'}) {
            assertThatThrownBy(() -> SheetOutput.target(DUMMY_FILE, "Sheet" + c + "Name"))
                    .as("char: " + c)
                    .isInstanceOf(IllegalArgumentException.class);
        }
    }

    @Test
    void validatesAcrossAllOverloads() {
        assertThatThrownBy(() -> SheetOutput.target(DUMMY_FILE, "Bad/Name"))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> SheetOutput.target(DUMMY_PATH, "Bad/Name"))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> SheetOutput.target(DUMMY_STREAM, "Bad/Name"))
                .isInstanceOf(IllegalArgumentException.class);
    }
}
