package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.SpreadsheetML;

/**
 * File-backed shared string table using H2 MVStore.
 * <p>
 * Stores shared strings on disk instead of in memory, keeping heap usage
 * constant regardless of shared string table size. Suitable for extremely
 * large spreadsheets where even the SoA-based {@link InMemorySharedStringLookup}
 * would consume too much memory.
 * <p>
 * Requires {@code com.h2database:h2} on the classpath.
 */
final class FileBackedSharedStringLookup implements SharedStringLookup {

    private static final Matcher START_SI = Matcher.startElement(SpreadsheetML.STRING_ITEM);

    private final XmlElementReader _reader;

    private final Path _storePath;
    private final MVStore _store;
    private final MVMap<Integer, String> _strings;
    private int _size;

    FileBackedSharedStringLookup(final PackagePart part) throws IOException {
        _reader = new XmlElementReader(part.getInputStream());
        _reader.navigateTo(SpreadsheetML.SST);

        _storePath = Files.createTempFile("jackson-sst-", ".mv");
        try {
            _store = new MVStore.Builder()
                    .fileName(_storePath.toString())
                    .cacheSize(4)
                    .autoCommitDisabled()
                    .open();
            _strings = _store.openMap("sharedStrings");
        } catch (RuntimeException e) {
            _reader.close();
            try { Files.deleteIfExists(_storePath); } catch (IOException ignored) {}
            throw e;
        }
    }

    @Override
    public String getItemAt(final int idx) {
        while (idx >= _size) {
            _reader.nextUntil(START_SI);
            final String text = _reader.readStringContent();
            _strings.put(_size, text != null ? text : "");
            _size++;
            if (_size % 10000 == 0) {
                _store.commit();
            }
        }
        return _strings.get(idx);
    }

    @Override
    public void close() throws IOException {
        _reader.close();
        _store.close();
        Files.deleteIfExists(_storePath);
    }
}
