package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;

import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.POICompat;

/**
 * File-backed shared string store using H2 MVStore.
 * Maintains constant heap usage for large/high-cardinality string sets.
 */
final class FileBackedSharedStringsStore implements SharedStringsStore {

    private static final int COMMIT_INTERVAL = 10000;

    private final Path _storePath;
    private final MVStore _store;
    private final MVMap<String, Integer> _indexByString;
    private final MVMap<Integer, String> _stringsByIndex;
    private int _size;
    private int _cachedIndex = -1;
    private String _cachedValue = "";
    private String _lastAddedValue = "";
    private int _lastAddedIndex = -1;

    FileBackedSharedStringsStore(final boolean encrypt) throws IOException {
        _storePath = _createSecureTempFile();
        try {
            MVStore.Builder builder = new MVStore.Builder()
                    .fileName(_storePath.toString())
                    .cacheSize(4)
                    .autoCommitDisabled();
            if (encrypt) {
                final char[] key = EncryptionKeys.generate();
                builder.encryptionKey(key);
                Arrays.fill(key, '\0');
            }
            _store = builder.open();
            _indexByString = _store.openMap("stringToIndex");
            _stringsByIndex = _store.openMap("indexToString");
        } catch (RuntimeException e) {
            try {
                POICompat.releaseTempFile(_storePath);
            } catch (IOException ignored) {
            }
            throw e;
        }
    }

    @Override
    public int getOrAdd(final String value) {
        final String s = value == null ? "" : value;
        if (s.equals(_lastAddedValue) && _lastAddedIndex >= 0) {
            return _lastAddedIndex;
        }
        final Integer existing = _indexByString.get(s);
        if (existing != null) {
            _lastAddedValue = s;
            _lastAddedIndex = existing;
            return existing;
        }
        final int index = _size;
        _indexByString.put(s, index);
        _stringsByIndex.put(index, s);
        _size = index + 1;
        _lastAddedValue = s;
        _lastAddedIndex = index;
        if (_size % COMMIT_INTERVAL == 0) {
            _store.commit();
        }
        return index;
    }

    @Override
    public int size() {
        return _size;
    }

    @Override
    public String get(final int index) {
        return _value(index);
    }

    @Override
    public int length(final int index) {
        return _value(index).length();
    }

    @Override
    public char charAt(final int index, final int offset) {
        return _value(index).charAt(offset);
    }

    @Override
    public void close() throws IOException {
        try {
            try {
                _store.commit();
            } finally {
                _store.close();
            }
        } finally {
            POICompat.releaseTempFile(_storePath);
        }
    }

    private static Path _createSecureTempFile() throws IOException {
        return POICompat.createSecureTempFile("jackson-spreadsheet-sst-write-", ".mv");
    }

    private String _value(final int index) {
        if (_cachedIndex != index) {
            final String value = _stringsByIndex.get(index);
            _cachedValue = value == null ? "" : value;
            _cachedIndex = index;
        }
        return _cachedValue;
    }
}
