package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.openxml4j.opc.PackagePart;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.SpreadsheetML;

/**
 * Lazy, streaming shared string table that reads entries on demand via StAX.
 * <p>
 * Uses Structure of Array (SoA) layout: all character data in a single {@code char[]} buffer
 * with separate {@code int[]} arrays for offsets and lengths.
 * Eliminates per-String object overhead (~56 bytes each).
 */
final class InMemorySharedStringLookup implements SharedStringLookup {

    private static final Matcher START_SI = Matcher.startElement(SpreadsheetML.STRING_ITEM);

    private static final int INITIAL_DATA_CAPACITY = 4096;

    private final XmlElementReader _reader;
    private final int _uniqueCount;

    private char[] _data;
    private int[] _offsets;
    private int[] _lengths;
    private int _size;
    private int _dataPos;

    InMemorySharedStringLookup(final PackagePart part) throws IOException {
        _reader = new XmlElementReader(part.getInputStream());
        _reader.navigateTo(SpreadsheetML.SST);
        final String uc = _reader.attribute(SpreadsheetML.ATTR_UNIQUE_COUNT);
        _uniqueCount = uc != null ? Integer.parseInt(uc) : 0;
        _offsets = new int[_uniqueCount];
        _lengths = new int[_uniqueCount];
        _data = new char[Math.max(INITIAL_DATA_CAPACITY, _uniqueCount * 16)];
    }

    @Override
    public String getItemAt(final int idx) {
        while (idx >= _size) {
            _reader.nextUntil(START_SI);
            final String text = _reader.readStringContent();
            _appendString(text != null ? text : "");
        }
        return new String(_data, _offsets[idx], _lengths[idx]);
    }

    private void _appendString(final String value) {
        final int len = value.length();
        _ensureDataCapacity(len);
        _ensureIndexCapacity();
        value.getChars(0, len, _data, _dataPos);
        _offsets[_size] = _dataPos;
        _lengths[_size] = len;
        _dataPos += len;
        _size++;
    }

    private void _ensureDataCapacity(final int additionalChars) {
        final int required = _dataPos + additionalChars;
        if (required > _data.length) {
            _data = Arrays.copyOf(_data, Math.max(required, _data.length * 2));
        }
    }

    private void _ensureIndexCapacity() {
        if (_size >= _offsets.length) {
            final int newCapacity = Math.max(_size + 1, _offsets.length * 2);
            _offsets = Arrays.copyOf(_offsets, newCapacity);
            _lengths = Arrays.copyOf(_lengths, newCapacity);
        }
    }

    @Override
    public void close() {
        _reader.close();
    }
}
