package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import com.fasterxml.jackson.core.JsonToken;
import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamReadException;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import support.FixtureAs;
import support.fixture.Entry;
import support.fixture.NestedEntry;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SheetParserTest implements FixtureAs {

    SpreadsheetMapper mapper;
    SheetParser parser;

    @BeforeEach
    void setUp() throws Exception {
        mapper = new SpreadsheetMapper();
        parser = mapper.createParser(fixtureAsFile("entries.xlsx"));
    }

    @Test
    void noSchema() throws Exception {
        assertThatThrownBy(parser::nextToken)
                .isInstanceOf(SheetStreamReadException.class)
                .hasMessageContaining("No schema of type '%s' set, can not parse", SpreadsheetSchema.SCHEMA_TYPE);
    }

    @Test
    void flatEntry() throws Exception {
        parser.setSchema(mapper.sheetSchemaFor(Entry.class));
        testEntry(this::Entry);
    }

    @Test
    void nestedEntry() throws Exception {
        parser.setSchema(mapper.sheetSchemaFor(NestedEntry.class));
        testEntry(this::assertNestedEntry);
    }

    void testEntry(final EntryAssertion assertEntry) throws Exception {
        assertThat(parser.hasCurrentToken()).isFalse();
        assertToken(JsonToken.START_ARRAY);
        assertEntry.run();
        assertEntry.run();
        assertToken(JsonToken.END_ARRAY);
        assertNoEntry();
        parser.close();
        assertThat(parser.isClosed()).isTrue();
    }

    void Entry() throws IOException {
        assertToken(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertToken(JsonToken.END_OBJECT);
    }

    void assertNestedEntry() throws IOException {
        assertToken(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertField(JsonToken.START_OBJECT);
        assertField(JsonToken.VALUE_NUMBER_INT);
        assertToken(JsonToken.END_OBJECT);
        assertToken(JsonToken.END_OBJECT);
    }

    void assertField(final JsonToken value) throws IOException {
        assertToken(JsonToken.FIELD_NAME);
        assertToken(value);
    }

    void assertNoEntry() throws IOException {
        assertThat(parser.nextToken()).isNull();
        assertThat(parser.hasCurrentToken()).isFalse();
    }

    void assertToken(final JsonToken token) throws IOException {
        assertThat(parser.nextToken()).isEqualTo(token);
        assertThat(parser.hasCurrentToken()).isTrue();
    }

    interface EntryAssertion {
        void run() throws IOException;
    }
}
