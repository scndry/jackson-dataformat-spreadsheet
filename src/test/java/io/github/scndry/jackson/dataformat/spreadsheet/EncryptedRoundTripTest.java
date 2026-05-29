package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.EncryptedDocumentException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class EncryptedRoundTripTest {

    @TempDir Path tempDir;

    @DataGrid
    @Data @NoArgsConstructor @AllArgsConstructor
    static class Employee {
        @DataColumn String name;
        @DataColumn int salary;
    }

    @Test
    void roundTrip_encryptedFile() throws Exception {
        final File file = tempDir.resolve("encrypted.xlsx").toFile();
        final List<Employee> input = Arrays.asList(
                new Employee("Alice", 80000),
                new Employee("Bob", 65000));

        final SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(
                SheetOutput.target(file).withPassword("secret"),
                input, Employee.class);

        assertThat(file).exists();
        assertThat(file.length()).isGreaterThan(0L);

        final List<Employee> output = mapper.readValues(
                SheetInput.source(file).withPassword("secret"),
                Employee.class);
        assertThat(output).isEqualTo(input);
    }

    @Test
    void read_wrongPassword_throws() throws Exception {
        final File file = tempDir.resolve("wrong-pwd.xlsx").toFile();
        final SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(
                SheetOutput.target(file).withPassword("secret"),
                Collections.singletonList(new Employee("A", 1)),
                Employee.class);

        assertThatThrownBy(() -> mapper.readValues(
                SheetInput.source(file).withPassword("wrong"),
                Employee.class))
                .isInstanceOf(EncryptedDocumentException.class)
                .hasMessageContaining("Invalid password");
    }

    @Test
    void read_plainFile_withPassword_throws() throws Exception {
        final File file = tempDir.resolve("plain.xlsx").toFile();
        final SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file,
                Collections.singletonList(new Employee("A", 1)),
                Employee.class);

        assertThatThrownBy(() -> mapper.readValues(
                SheetInput.source(file).withPassword("secret"),
                Employee.class))
                .isInstanceOf(EncryptedDocumentException.class)
                .hasMessageContaining("not an encrypted OOXML document");
    }

    @Test
    void roundTrip_namedSheet() throws Exception {
        final File file = tempDir.resolve("named.xlsx").toFile();
        final List<Employee> input = Collections.singletonList(new Employee("Alice", 80000));

        final SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(
                SheetOutput.target(file, "People").withPassword("secret"),
                input, Employee.class);

        final List<Employee> output = mapper.readValues(
                SheetInput.source(file, "People").withPassword("secret"),
                Employee.class);
        assertThat(output).isEqualTo(input);
    }

    @Test
    void roundTrip_poiUserModel() throws Exception {
        final File file = tempDir.resolve("poi-um.xlsx").toFile();
        final List<Employee> input = SAMPLE;
        final SpreadsheetMapper mapper = poiMapper();

        mapper.writeValue(
                SheetOutput.target(file).withPassword("secret"),
                input, Employee.class);
        final List<Employee> output = mapper.readValues(
                SheetInput.source(file).withPassword("secret"),
                Employee.class);
        assertThat(output).isEqualTo(input);
    }

    @Test
    void roundTrip_streamRaw() throws Exception {
        final File file = tempDir.resolve("stream.xlsx").toFile();
        final List<Employee> input = SAMPLE;
        final SpreadsheetMapper mapper = new SpreadsheetMapper();

        try (OutputStream out = Files.newOutputStream(file.toPath())) {
            mapper.writeValue(
                    SheetOutput.target(out).withPassword("secret"),
                    input, Employee.class);
        }
        try (InputStream in = Files.newInputStream(file.toPath())) {
            final List<Employee> output = mapper.readValues(
                    SheetInput.source(in).withPassword("secret"),
                    Employee.class);
            assertThat(output).isEqualTo(input);
        }
    }

    @Test
    void crossPath_ssmlWriteEncrypt_poiReadDecrypt() throws Exception {
        final File file = tempDir.resolve("cross-sw-pr.xlsx").toFile();
        final List<Employee> input = SAMPLE;

        new SpreadsheetMapper().writeValue(
                SheetOutput.target(file).withPassword("secret"),
                input, Employee.class);
        final List<Employee> output = poiMapper().readValues(
                SheetInput.source(file).withPassword("secret"),
                Employee.class);
        assertThat(output).isEqualTo(input);
    }

    @Test
    void crossPath_poiWriteEncrypt_ssmlReadDecrypt() throws Exception {
        final File file = tempDir.resolve("cross-pw-sr.xlsx").toFile();
        final List<Employee> input = SAMPLE;

        poiMapper().writeValue(
                SheetOutput.target(file).withPassword("secret"),
                input, Employee.class);
        final List<Employee> output = new SpreadsheetMapper().readValues(
                SheetInput.source(file).withPassword("secret"),
                Employee.class);
        assertThat(output).isEqualTo(input);
    }

    private static SpreadsheetMapper poiMapper() {
        return SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
    }

    private static final List<Employee> SAMPLE = Arrays.asList(
            new Employee("Alice", 80000),
            new Employee("Bob", 65000));
}
