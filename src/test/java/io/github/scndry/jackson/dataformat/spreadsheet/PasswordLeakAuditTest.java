package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;

import com.fasterxml.jackson.databind.ObjectMapper;

import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * SEC-21 — Jackson serialize 의 password leak 검증.
 */
class PasswordLeakAuditTest {

    private static final String SECRET = "MY_VERY_SECRET_PASSWORD";

    @Test
    void sheetInput_objectMapper_writeValueAsString_password_leak() throws Exception {
        final ObjectMapper m = new ObjectMapper();
        final SheetInput<File> in = SheetInput.source(new File("/tmp/x.xlsx")).withPassword(SECRET);
        final String json = m.writeValueAsString(in);
        // SEC-21 fail 검출: password 가 JSON 에 포함되면 leak
        assertThat(json).as("SheetInput JSON 에 password 가 포함 (SEC-21)").doesNotContain(SECRET);
    }

    @Test
    void sheetOutput_objectMapper_writeValueAsString_password_leak() throws Exception {
        final ObjectMapper m = new ObjectMapper();
        final SheetOutput<File> out = SheetOutput.target(new File("/tmp/y.xlsx")).withPassword(SECRET);
        final String json = m.writeValueAsString(out);
        assertThat(json).as("SheetOutput JSON 에 password 가 포함 (SEC-21)").doesNotContain(SECRET);
    }

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
