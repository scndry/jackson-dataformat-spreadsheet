package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.h2.mvstore.MVStore;
import org.h2.mvstore.MVStoreException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Verifies {@code ENCRYPT_FILE_BACKED_STORE} is actually applied to the
 * intermediate H2 MVStore on both the write side ({@code FileBackedSharedStringsStore})
 * and the read side ({@code FileBackedSharedStringsLookup}).
 * <p>
 * The existing round-trip tests ({@code FileBackedSharedStringsTest},
 * {@code SSMLSheetWriterIntegrationTest}) exercise outcome correctness only —
 * they would pass even if encryption silently became a no-op, because the
 * final XLSX is unchanged. A silent regression would mean the temp store
 * containing user data is plaintext-recoverable from disk, contradicting
 * the library's encryption claim.
 */
class FileBackedSharedStringsStoreEncryptionTest {

    private static final String SECRET_MARKER = "sensitive-payload-marker-xyz";
    private static final char[] WRONG_KEY = "wrong-key-deadbeef".toCharArray();

    @TempDir
    Path tempDir;

    // -- Store (write side) --

    @Test
    void store_encryptedFile_cannotBeOpenedWithoutKeyOrWithWrongKey() throws Exception {
        Path retainedPath = buildAndRetainStoreFile(true);
        try {
            assertThatThrownBy(() ->
                    new MVStore.Builder().fileName(retainedPath.toString()).open())
                    .isInstanceOf(MVStoreException.class);
            assertThatThrownBy(() ->
                    new MVStore.Builder().fileName(retainedPath.toString())
                            .encryptionKey(WRONG_KEY.clone()).open())
                    .isInstanceOf(MVStoreException.class);
        } finally {
            Files.deleteIfExists(retainedPath);
        }
    }

    @Test
    void store_encryptedFile_doesNotContainPlaintext() throws Exception {
        Path retainedPath = buildAndRetainStoreFile(true);
        try {
            assertThat(readUtf8(retainedPath)).doesNotContain(SECRET_MARKER);
        } finally {
            Files.deleteIfExists(retainedPath);
        }
    }

    /**
     * Confirms the marker IS visible in the unencrypted file format,
     * so the marker-absence assertion above is meaningful — not a false
     * positive from H2 storing strings encoded regardless of encryption.
     */
    @Test
    void store_unencryptedFile_containsPlaintext() throws Exception {
        Path retainedPath = buildAndRetainStoreFile(false);
        try {
            assertThat(readUtf8(retainedPath)).contains(SECRET_MARKER);
        } finally {
            Files.deleteIfExists(retainedPath);
        }
    }

    // -- Lookup (read side) --

    @Test
    void lookup_encryptedFile_cannotBeOpenedWithoutKeyOrWithWrongKey() throws Exception {
        Path retainedPath = buildAndRetainLookupFile(true);
        try {
            assertThatThrownBy(() ->
                    new MVStore.Builder().fileName(retainedPath.toString()).open())
                    .isInstanceOf(MVStoreException.class);
            assertThatThrownBy(() ->
                    new MVStore.Builder().fileName(retainedPath.toString())
                            .encryptionKey(WRONG_KEY.clone()).open())
                    .isInstanceOf(MVStoreException.class);
        } finally {
            Files.deleteIfExists(retainedPath);
        }
    }

    @Test
    void lookup_encryptedFile_doesNotContainPlaintext() throws Exception {
        Path retainedPath = buildAndRetainLookupFile(true);
        try {
            assertThat(readUtf8(retainedPath)).doesNotContain(SECRET_MARKER);
        } finally {
            Files.deleteIfExists(retainedPath);
        }
    }

    @Test
    void lookup_unencryptedFile_containsPlaintext() throws Exception {
        Path retainedPath = buildAndRetainLookupFile(false);
        try {
            assertThat(readUtf8(retainedPath)).contains(SECRET_MARKER);
        } finally {
            Files.deleteIfExists(retainedPath);
        }
    }

    // -- Helpers --

    /**
     * Builds the write-side store, writes the marker, then closes the underlying
     * MVStore directly via reflection — releasing the file lock without deleting
     * the temp file (the public close() would delete it). Caller deletes the path.
     */
    private static Path buildAndRetainStoreFile(final boolean encrypt) throws Exception {
        final FileBackedSharedStringsStore store = new FileBackedSharedStringsStore(encrypt);
        try {
            final Path storePath = readField(store, "_storePath", Path.class);
            store.getOrAdd(SECRET_MARKER);
            final MVStore mvStore = readField(store, "_store", MVStore.class);
            mvStore.commit();
            mvStore.close();
            return storePath;
        } catch (Exception e) {
            try { store.close(); } catch (Exception cleanup) { e.addSuppressed(cleanup); }
            throw e;
        }
    }

    /**
     * Builds an XLSX whose shared strings table contains the marker, opens it
     * as a Lookup, triggers storage of the marker into the H2 MVStore, then
     * closes the MVStore and reader directly — preserving the temp file for
     * inspection.
     */
    private Path buildAndRetainLookupFile(final boolean encrypt) throws Exception {
        final File xlsx = tempDir.resolve("encryption-lookup.xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0).createCell(0).setCellValue("header_unique");
            sheet.createRow(1).createCell(0).setCellValue(SECRET_MARKER);
            try (FileOutputStream os = new FileOutputStream(xlsx)) {
                wb.write(os);
            }
        }

        try (SSMLWorkbook workbook = SSMLWorkbook.create(xlsx)) {
            final PackagePart sst = workbook.getSharedStringsPart();
            assertThat(sst).isNotNull();

            final FileBackedSharedStringsLookup lookup = new FileBackedSharedStringsLookup(sst, encrypt);
            try {
                final Path storePath = readField(lookup, "_storePath", Path.class);
                // getItemAt iterates from 0 and stores each as it goes — index 1
                // forces the marker to be persisted to the MVStore.
                lookup.getItemAt(1);
                final XmlElementReader reader = readField(lookup, "_reader", XmlElementReader.class);
                final MVStore mvStore = readField(lookup, "_store", MVStore.class);
                mvStore.commit();
                mvStore.close();
                reader.close();
                return storePath;
            } catch (Exception e) {
                try { lookup.close(); } catch (Exception cleanup) { e.addSuppressed(cleanup); }
                throw e;
            }
        }
    }

    private static String readUtf8(final Path path) throws Exception {
        return new String(Files.readAllBytes(path), StandardCharsets.UTF_8);
    }

    private static <T> T readField(final Object target, final String name, final Class<T> type) throws Exception {
        final Field field = target.getClass().getDeclaredField(name);
        field.setAccessible(true);
        return type.cast(field.get(target));
    }
}
