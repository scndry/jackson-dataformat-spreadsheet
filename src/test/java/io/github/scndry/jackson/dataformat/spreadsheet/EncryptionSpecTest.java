package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class EncryptionSpecTest {

    @TempDir Path tempDir;

    @DataGrid
    @Data @NoArgsConstructor @AllArgsConstructor
    static class Row {
        @DataColumn String name;
        @DataColumn int qty;
    }

    private static final List<Row> SAMPLE = Arrays.asList(new Row("A", 1), new Row("B", 2));

    private void _roundTrip(final EncryptionSpec spec, final String fileName) throws Exception {
        final File file = tempDir.resolve(fileName).toFile();
        final SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(
                SheetOutput.target(file).withPassword("secret", spec),
                SAMPLE, Row.class);
        final List<Row> back = mapper.readValues(
                SheetInput.source(file).withPassword("secret"), Row.class);
        assertThat(back).isEqualTo(SAMPLE);
    }

    @Test
    void preset_strong_roundTrip() throws Exception {
        _roundTrip(EncryptionSpec.strong(), "strong.xlsx");
    }

    @Test
    void preset_balanced_roundTrip() throws Exception {
        _roundTrip(EncryptionSpec.balanced(), "balanced.xlsx");
    }

    @Test
    void preset_legacy_roundTrip() throws Exception {
        _roundTrip(EncryptionSpec.legacy(), "legacy.xlsx");
    }

    @Test
    void preset_fast_roundTrip() throws Exception {
        _roundTrip(EncryptionSpec.fast(), "fast.xlsx");
    }

    @Test
    void custom_aes192_sha384_roundTrip() throws Exception {
        _roundTrip(EncryptionSpec.custom()
                .cipher(EncryptionSpec.Cipher.AES_192)
                .hash(EncryptionSpec.Hash.SHA_384)
                .build(), "custom-aes192.xlsx");
    }

    @Test
    void default_strong_whenNoExplicitSpec() throws Exception {
        final File file = tempDir.resolve("default.xlsx").toFile();
        final SpreadsheetMapper mapper = new SpreadsheetMapper();
        // No withEncryption() → defaults to strong()
        mapper.writeValue(
                SheetOutput.target(file).withPassword("secret"),
                SAMPLE, Row.class);
        final List<Row> back = mapper.readValues(
                SheetInput.source(file).withPassword("secret"), Row.class);
        assertThat(back).isEqualTo(SAMPLE);
    }

    @Test
    void custom_invalidCombo_aes256OnLegacy_throws() {
        assertThatThrownBy(() -> EncryptionSpec.custom()
                .compatibility(EncryptionSpec.Compatibility.EXCEL_2007_2010)
                .cipher(EncryptionSpec.Cipher.AES_256)
                .build())
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("does not support AES-256");
    }

    @Test
    void custom_invalidKeyBits_throws() {
        assertThatThrownBy(() -> EncryptionSpec.custom().keyBits(160))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("keyBits must be 128, 192, or 256");
    }
}
