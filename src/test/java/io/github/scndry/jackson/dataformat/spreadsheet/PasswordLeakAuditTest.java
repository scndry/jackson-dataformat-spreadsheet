package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;

import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;

import static org.assertj.core.api.Assertions.assertThat;

class PasswordLeakAuditTest {

    private static final String SECRET = "MY_VERY_SECRET_PASSWORD";

    @Test
    void sheetInput_toString_password_masked() {
        final SheetInput<File> in = SheetInput.source(new File("/tmp/x.xlsx")).withPassword(SECRET);
        final String s = in.toString();
        assertThat(s).doesNotContain(SECRET);
        assertThat(s).contains("password=***");
    }

    @Test
    void sheetOutput_toString_password_masked() {
        final SheetOutput<File> out = SheetOutput.target(new File("/tmp/y.xlsx")).withPassword(SECRET);
        final String s = out.toString();
        assertThat(s).doesNotContain(SECRET);
        assertThat(s).contains("password=***");
    }
}
