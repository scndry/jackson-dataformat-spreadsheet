package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
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
}
