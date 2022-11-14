package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import support.FixtureAs;
import support.fixture.Entry;

import java.io.File;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SpreadsheetMapperTest implements FixtureAs {

    SpreadsheetMapper mapper;

    @BeforeEach
    void setUp() throws Exception {
        mapper = new SpreadsheetMapper();
    }

    @Nested
    class ReadTest {
        @Test
        void readValue() throws Exception {
            final File src = fixtureAsFile("entries.xlsx");
            final Entry value = mapper.readValue(src, Entry.class);
            assertThat(value).isEqualTo(Entry.VALUE);
        }

        @Test
        void readValues() throws Exception {
            final File src = fixtureAsFile("entries.xlsx");
            final List<Entry> value = mapper.readValues(src, Entry.class);
            assertThat(value).isEqualTo(Entry.VALUES);
        }

        @Test
        void readValuesWithSheetName() throws Exception {
            final SheetInput<File> src = SheetInput.source(fixtureAsFile("entries.xlsx"), "Entries");
            final List<Entry> value = mapper.readValues(src, Entry.class);
            assertThat(value).isEqualTo(Entry.VALUES);
        }
    }

    @Nested
    class ReadFailingTest {
        @Test
        void readValuesNoSheet() throws Exception {
            final SheetInput<File> src = SheetInput.source(fixtureAsFile("entries.xlsx"), "NoSheet");
            assertThatThrownBy(() -> mapper.readValues(src, Entry.class))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageStartingWith("No sheet for");
        }

        @Test
        void readValuesOutOfRange() throws Exception {
            final SheetInput<File> src = SheetInput.source(fixtureAsFile("entries.xlsx"), 1);
            assertThatThrownBy(() -> mapper.readValues(src, Entry.class))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("out of range");
        }
    }

    @Nested
    class WriteTest {
        @TempDir
        Path tempDir;
        File out;

        @BeforeEach
        void setUp() throws Exception {
            out = tempDir.resolve("test.xlsx").toFile();
        }

        @Test
        void writeValue() throws Exception {
            mapper.writeValue(out, Entry.VALUE);
            final Entry actual = mapper.readValue(out, Entry.class);
            assertThat(actual).isEqualTo(Entry.VALUE).isNotSameAs(Entry.VALUE);
        }

        @Test
        void writeValues() throws Exception {
            mapper.writeValue(out, Entry.VALUES, Entry.class);
            final List<Entry> actual = mapper.readValues(out, Entry.class);
            assertThat(actual).isEqualTo(Entry.VALUES).isNotSameAs(Entry.VALUES);
        }

        @Test
        void writeValuesWithSheetName() throws Exception {
            final SheetOutput<File> output = SheetOutput.target(out, "Entries");
            mapper.writeValue(output, Entry.VALUES, Entry.class);
            try (XSSFWorkbook workbook = new XSSFWorkbook(out)) {
                final XSSFSheet sheet = workbook.getSheet("Entries");
                final List<Entry> actual = mapper.readValues(sheet, Entry.class);
                assertThat(actual).isEqualTo(Entry.VALUES).isNotSameAs(Entry.VALUES);
            }
        }
    }

    @Nested
    class WriteFailingTest {

        @TempDir
        Path tempDir;
        File out;

        @BeforeEach
        void setUp() throws Exception {
            out = tempDir.resolve("test.xlsx").toFile();
        }

        @Test
        void writeValuesWithoutValueType() throws Exception {
            assertThatThrownBy(() -> mapper.writeValue(out, Entry.VALUES))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessage("`valueType` MUST be specified to write a value of a Collection or array type");
        }
    }
}
