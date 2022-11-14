package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import support.FixtureAs;

import java.io.IOException;
import java.io.InputStream;
import java.util.NoSuchElementException;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

public class SheetReaderTestBase implements FixtureAs {

    protected final InputStream source = fixtureAsStream("entries-headless.xlsx");
    protected SheetReader reader;

    protected void testSheetReader() throws Exception {
        assertNext(SheetToken.SHEET_DATA_START);
        assertNoRow();
        assertNoCellValue();
        for (int i = 0; i < 2; i++) {
            assertNext(SheetToken.ROW_START);
            assertRow();
            assertNoCellValue();
            for (int j = 0; j < 2; j++) {
                assertNext(SheetToken.CELL_VALUE);
                assertRow();
                assertCellValue();
            }
            assertNext(SheetToken.ROW_END);
            assertRow();
            assertNoCellValue();
        }
        assertNext(SheetToken.SHEET_DATA_END);
        assertExhausted();
        assertClose();
    }

    void assertNext(final SheetToken token) {
        assertThat(reader).hasNext();
        assertThat(reader.next()).isEqualTo(token);
    }

    void assertRow() {
        assertThat(reader.getRow()).isNotNegative();
    }

    void assertNoRow() {
        assertThat(reader.getRow()).isNegative();
    }

    void assertCellValue() {
        assertThat(reader.getReference()).isNotNull();
        assertThat(reader.getCellValue()).isNotNull();
    }

    void assertNoCellValue() {
        assertThat(reader.getReference()).isNull();
        assertThat(reader.getCellValue()).isNull();
    }

    void assertExhausted() {
        assertThat(reader).isExhausted();
        assertThatThrownBy(reader::next).isInstanceOf(NoSuchElementException.class);
    }

    void assertClose() throws IOException {
        reader.close();
        assertThat(reader.isClosed()).isTrue();
    }
}
