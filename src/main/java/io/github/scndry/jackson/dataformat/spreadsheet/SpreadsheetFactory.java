package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;

import java.security.GeneralSecurityException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.crypt.temp.EncryptedTempData;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.extern.slf4j.Slf4j;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.io.IOContext;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.POICompat;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.PackageUtil;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.SSMLSheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.SSMLSheetWriter;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.SSMLWorkbook;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ss.POISheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ss.POISheetWriter;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetGenerator;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetWriter;

/**
 * {@link JsonFactory} implementation for spreadsheet formats.
 * Creates {@link SheetParser} and {@link SheetGenerator} instances
 * for reading and writing Excel files.
 *
 * @see SpreadsheetMapper
 * @see SheetParser
 * @see SheetGenerator
 */
@Slf4j
@SuppressWarnings("java:S2177")
public final class SpreadsheetFactory extends JsonFactory {

    public static final String FORMAT_NAME = "spreadsheet";
    public static final int DEFAULT_SHEET_PARSER_FEATURE_FLAGS =
            SheetParser.Feature.collectDefaults();
    public static final int DEFAULT_FEATURE_FLAGS = Feature.collectDefaults();

    private final transient WorkbookProvider _workbookProvider;
    private int _sheetParserFeatures;
    private int _featureFlags;

    public SpreadsheetFactory() {
        this(XSSFWorkbook::new, DEFAULT_SHEET_PARSER_FEATURE_FLAGS, DEFAULT_FEATURE_FLAGS);
    }

    public SpreadsheetFactory(
            final WorkbookProvider workbookProvider,
            final int sheetParserFeatures) {
        this(workbookProvider, sheetParserFeatures, DEFAULT_FEATURE_FLAGS);
    }

    public SpreadsheetFactory(
            final WorkbookProvider workbookProvider,
            final int sheetParserFeatures,
            final int featureFlags) {
        _workbookProvider = workbookProvider;
        _sheetParserFeatures = sheetParserFeatures;
        _featureFlags = featureFlags;
    }

    public SpreadsheetFactory(final SpreadsheetFactory base) {
        _workbookProvider = base._workbookProvider;
        _sheetParserFeatures = base._sheetParserFeatures;
        _featureFlags = base._featureFlags;
    }

    @Override
    public SpreadsheetFactory copy() {
        _checkInvalidCopy(SpreadsheetFactory.class);
        return new SpreadsheetFactory(this);
    }

    /*
    /**********************************************************
    /* Format detection functionality
    /**********************************************************
     */

    @Override
    public boolean canUseSchema(final FormatSchema schema) {
        return schema instanceof SpreadsheetSchema;
    }

    @Override
    public String getFormatName() {
        return FORMAT_NAME;
    }

    /*
    /**********************************************************
    /* Versioned
    /**********************************************************
     */

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    /*
    /**********************************************************
    /* Configuration, sheet parser configuration
    /**********************************************************
     */

    public SpreadsheetFactory configure(final SheetParser.Feature f, final boolean state) {
        return state ? enable(f) : disable(f);
    }

    public SpreadsheetFactory enable(final SheetParser.Feature f) {
        _sheetParserFeatures |= f.getMask();
        return this;
    }

    public SpreadsheetFactory disable(final SheetParser.Feature f) {
        _sheetParserFeatures &= ~f.getMask();
        return this;
    }

    /*
    /**********************************************************
    /* Configuration, factory feature configuration
    /**********************************************************
     */

    public SpreadsheetFactory configure(final Feature f, final boolean state) {
        return state ? enable(f) : disable(f);
    }

    public SpreadsheetFactory enable(final Feature f) {
        _featureFlags |= f.getMask();
        return this;
    }

    public SpreadsheetFactory disable(final Feature f) {
        _featureFlags &= ~f.getMask();
        return this;
    }

    /*
    /**********************************************************
    /* Parser factories
    /**********************************************************
     */

    public SheetParser createParser(final Sheet src) {
        final IOContext ctxt = _createContext(_createContentReference(src), false);
        return _createParser(new POISheetReader(src), ctxt);
    }

    @SuppressWarnings("unchecked")
    public SheetParser createParser(final SheetInput<?> src) throws IOException {
        SheetInput<?> source = src;
        if (src.getPassword() != null) {
            source = _decryptToSheetInput(src);
        }
        source = _preferRawAsFile(source);
        final boolean resourceManaged = src != source;
        final IOContext ctxt = _createContext(_createContentReference(source), resourceManaged);
        final SheetReader reader;
        if (source.isFile()) {
            reader = _createFileSheetReader((SheetInput<File>) source);
        } else {
            reader = _createInputStreamSheetReader((SheetInput<InputStream>) source);
        }
        return _createParser(reader, ctxt);
    }

    @Override
    public SheetParser createParser(final File src) throws IOException {
        return createParser(SheetInput.source(src));
    }

    @Override
    public SheetParser createParser(final InputStream src) throws IOException {
        return createParser(SheetInput.source(src));
    }

    /*
    /**********************************************************
    /* Generator factories
    /**********************************************************
     */

    public SheetGenerator createGenerator(final Sheet out) {
        final IOContext ctxt = _createContext(_createContentReference(out), false);
        final SheetGenerator g = _createGenerator(new POISheetWriter(out), ctxt);
        g.disable(JsonGenerator.Feature.AUTO_CLOSE_TARGET);
        return g;
    }

    public SheetGenerator createGenerator(final SheetOutput<?> out) throws IOException {
        final SheetOutput<OutputStream> output = _rawAsOutputStream(out);
        final boolean resourceManaged = out != output;
        final IOContext ctxt = _createContext(_createContentReference(output), resourceManaged);
        return _createGenerator(_createSheetWriter(output), ctxt);
    }

    @Override
    public SheetGenerator createGenerator(
            final File out,
            final JsonEncoding enc) throws IOException {
        return createGenerator(SheetOutput.target(out));
    }

    @Override
    public SheetGenerator createGenerator(
            final OutputStream out,
            final JsonEncoding enc) throws IOException {
        return createGenerator(SheetOutput.target(out));
    }

    @Override
    public SheetGenerator createGenerator(final OutputStream out) throws IOException {
        return createGenerator(SheetOutput.target(out));
    }

    /*
    /**********************************************************
    /* Factory methods used by factory for creating parser instances,
    /**********************************************************
     */

    private SheetParser _createParser(final SheetReader reader, final IOContext ctxt) {
        return new SheetParser(ctxt, _parserFeatures, _objectCodec, _sheetParserFeatures, reader);
    }

    private SheetReader _createFileSheetReader(final SheetInput<File> src) throws IOException {
        if (_shouldUsePOIUserModel(src.getRaw())) {
            return _createPOISheetReader(WorkbookFactory.create(src.getRaw()), src);
        }
        return _createSSMLSheetReader(SSMLWorkbook.create(src.getRaw()), src);
    }

    private SheetReader _createInputStreamSheetReader(
            final SheetInput<InputStream> src) throws IOException {
        // Reachable only when _preferRawAsFile did not spool to a file:
        //   USE_POI_USER_MODEL enabled, or non-OOXML input.
        // Either way the read uses POI's User Model — SSML needs seekable I/O.
        return _createPOISheetReader(WorkbookFactory.create(src.getRaw()), src);
    }

    private SSMLSheetReader _createSSMLSheetReader(
            final SSMLWorkbook workbook,
            final SheetInput<?> src) {
        final PackagePart worksheetPart = src
                .isNamed() ? workbook
                .getWorksheetPart(src.getName()) : workbook
                .getWorksheetPartAt(src.getIndex());
        if (worksheetPart == null) {
            throw new IllegalArgumentException("No sheet for " + src);
        }
        final boolean fileBacked = Feature
                .FILE_BACKED_SHARED_STRINGS
                .enabledIn(_featureFlags);
        final boolean encrypt = Feature
                .ENCRYPT_FILE_BACKED_STORE
                .enabledIn(_featureFlags);
        return new SSMLSheetReader(worksheetPart, workbook, fileBacked, encrypt);
    }

    private POISheetReader _createPOISheetReader(final Workbook workbook, final SheetInput<?> src) {
        final Sheet sheet = src.isNamed()
                ? workbook.getSheet(src.getName()) : workbook.getSheetAt(src.getIndex());
        if (sheet == null) {
            throw new IllegalArgumentException("No sheet for " + src);
        }
        return new POISheetReader(sheet);
    }

    private boolean _shouldUsePOIUserModel(final File src) {
        return Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags)
            || !PackageUtil.isOOXML(src);
    }

    private boolean _shouldUsePOIUserModel(final InputStream src) {
        return Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags)
            || !PackageUtil.isOOXML(src);
    }

    // Copy InputStream to temp File — POI's File path uses much less heap than InputStream (POIFS HOWTO ~20% vs ~120% for HSSF; OPCPackage Javadoc reports the same trend for OOXML).
    @SuppressWarnings("unchecked")
    private SheetInput<?> _preferRawAsFile(final SheetInput<?> src) throws IOException {
        if (src.isFile()) return src;
        final InputStream raw = FileMagic.prepareToCheckMagic(
                ((SheetInput<InputStream>) src).getRaw());
        if (_shouldUsePOIUserModel(raw)) {
            return src.isNamed()
                    ? SheetInput.source(raw, src.getName())
                    : SheetInput.source(raw, src.getIndex());
        }
        File file = null;
        try {
            file = POICompat.createSecureTempFile("jackson-spreadsheet-input-", ".xlsx").toFile();
            try (FileOutputStream out = new FileOutputStream(file)) {
                final byte[] buf = new byte[8192];
                int n;
                while ((n = raw.read(buf)) != -1) {
                    out.write(buf, 0, n);
                }
            }
        } catch (IOException e) {
            if (file != null) {
                try { POICompat.releaseTempFile(file.toPath()); } catch (IOException cleanup) { e.addSuppressed(cleanup); }
            }
            throw new IOException(
                    "Failed to spool InputStream to temp file. "
                            + "If running in a disk-write-restricted environment "
                            + "(e.g. Lambda read-only filesystem, Kubernetes readOnlyRootFilesystem), "
                            + "enable SpreadsheetFactory.Feature.USE_POI_USER_MODEL to read directly "
                            + "from the InputStream (uses more heap; see POI WorkbookFactory javadoc).", e);
        }
        if (log.isDebugEnabled()) {
            log.debug("Copied InputStream to temp file: {}", file);
        }
        if (isEnabled(StreamReadFeature.AUTO_CLOSE_SOURCE)) {
            raw.close();
        }
        return src.isNamed()
                ? SheetInput.source(file, src.getName()) : SheetInput.source(file, src.getIndex());
    }

    /**
     * Decrypts an encrypted OOXML source to a POSIX owner-only temp file and
     * returns a fresh {@link SheetInput} referencing it. The temp file is
     * scheduled for deletion when the resulting {@link SheetParser} closes
     * (resource-managed). Caller-owned {@link InputStream} sources are first
     * spooled to disk to keep heap usage bounded for large encrypted files.
     *
     * <p>Throws {@link EncryptedDocumentException} when the password is wrong
     * or the source is not actually encrypted.
     */
    @SuppressWarnings("unchecked")
    private SheetInput<?> _decryptToSheetInput(final SheetInput<?> src) throws IOException {
        // F6: spool InputStream to disk before opening POIFS to avoid loading the
        // entire encrypted file into memory.
        final File encryptedFile;
        final boolean spooledFromStream;
        if (src.isFile()) {
            encryptedFile = ((SheetInput<File>) src).getRaw();
            spooledFromStream = false;
        } else {
            encryptedFile = _spoolToSecureTempFile(
                    ((SheetInput<InputStream>) src).getRaw(),
                    "jackson-spreadsheet-encrypted-");
            spooledFromStream = true;
        }

        File decrypted = null;
        try (POIFSFileSystem fs = _openEncryptedPOIFS(encryptedFile, src)) {
            final EncryptionInfo info;
            try {
                info = new EncryptionInfo(fs);
            } catch (IOException e) {
                throw new EncryptedDocumentException(
                        "Source is not an encrypted OOXML document (password was supplied"
                                + " but no EncryptionInfo entry found): " + src);
            }
            final Decryptor d = Decryptor.getInstance(info);
            try {
                if (!d.verifyPassword(src.getPassword())) {
                    throw new EncryptedDocumentException(
                            "Invalid password for encrypted spreadsheet source: " + src);
                }
            } catch (GeneralSecurityException e) {
                throw new IOException("Failed to verify password for " + src, e);
            }
            decrypted = POICompat.createSecureTempFile(
                    "jackson-spreadsheet-decrypted-", ".xlsx").toFile();
            try (InputStream plain = d.getDataStream(fs);
                 FileOutputStream out = new FileOutputStream(decrypted)) {
                final byte[] buf = new byte[8192];
                int n;
                while ((n = plain.read(buf)) != -1) {
                    out.write(buf, 0, n);
                }
            } catch (GeneralSecurityException e) {
                throw new IOException("Failed to read decrypted data stream for " + src, e);
            }
        } catch (IOException | RuntimeException e) {
            if (decrypted != null) {
                try { POICompat.releaseTempFile(decrypted.toPath()); }
                catch (IOException cleanup) { e.addSuppressed(cleanup); }
            }
            throw e;
        } finally {
            if (spooledFromStream) {
                try { POICompat.releaseTempFile(encryptedFile.toPath()); }
                catch (IOException ignore) { /* best effort */ }
            }
        }
        if (log.isDebugEnabled()) {
            log.debug("Decrypted encrypted source to temp file: {}", decrypted);
        }
        return src.isNamed()
                ? SheetInput.source(decrypted, src.getName())
                : SheetInput.source(decrypted, src.getIndex());
    }

    private POIFSFileSystem _openEncryptedPOIFS(final File encryptedFile, final SheetInput<?> src)
            throws IOException {
        try {
            return new POIFSFileSystem(encryptedFile, true);
        } catch (org.apache.poi.poifs.filesystem.OfficeXmlFileException e) {
            // F8: plain OOXML (zip) — POIFSFileSystem only opens OLE2 binary.
            throw new EncryptedDocumentException(
                    "Source is not an encrypted OOXML document (password was supplied"
                            + " but the file is a plain OOXML package): " + src);
        }
    }

    private File _spoolToSecureTempFile(final InputStream in, final String prefix) throws IOException {
        final File temp = POICompat.createSecureTempFile(prefix, ".tmp").toFile();
        try (FileOutputStream out = new FileOutputStream(temp)) {
            final byte[] buf = new byte[8192];
            int n;
            while ((n = in.read(buf)) != -1) {
                out.write(buf, 0, n);
            }
        } catch (IOException e) {
            try { POICompat.releaseTempFile(temp.toPath()); }
            catch (IOException cleanup) { e.addSuppressed(cleanup); }
            throw e;
        }
        return temp;
    }

    /*
    /**********************************************************
    /* Factory methods used by factory for creating generator instances,
    /**********************************************************
     */

    private SheetGenerator _createGenerator(final SheetWriter writer, final IOContext ctxt) {
        return new SheetGenerator(ctxt, _generatorFeatures, _objectCodec, writer);
    }

    @SuppressWarnings("resource")
    private SheetWriter _createSheetWriter(final SheetOutput<OutputStream> out) throws IOException {
        final Workbook workbook = _workbookProvider.create();
        final Sheet sheet = out.isNamed()
                ? workbook.createSheet(out.getName()) : workbook.createSheet();
        final boolean useUserModel = _shouldUsePOIUserModel(sheet);
        final boolean autoFallback = useUserModel && !Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags);
        if (autoFallback && log.isDebugEnabled()) {
            log.debug("Sheet is not XSSFSheet ({}); falling back to POISheetWriter",
                    sheet.getClass().getSimpleName());
        }
        if (useUserModel) {
            return _createPOISheetWriter(sheet, out.getRaw());
        }
        return _createSSMLSheetWriter(out.getRaw(), sheet);
    }

    private SSMLSheetWriter _createSSMLSheetWriter(
            final OutputStream out, final Sheet sheet) {
        final boolean fileBacked = Feature
                .FILE_BACKED_SHARED_STRINGS
                .enabledIn(_featureFlags);
        final boolean encrypt = Feature
                .ENCRYPT_FILE_BACKED_STORE
                .enabledIn(_featureFlags);
        return new SSMLSheetWriter(out, sheet, fileBacked, encrypt);
    }

    private POISheetWriter _createPOISheetWriter(final Sheet sheet, final OutputStream out) {
        return new POISheetWriter(sheet, out);
    }

    private boolean _shouldUsePOIUserModel(final Sheet sheet) {
        return Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags)
            || !(sheet instanceof XSSFSheet);
    }

    @SuppressWarnings("unchecked")
    private SheetOutput<OutputStream> _rawAsOutputStream(
            final SheetOutput<?> out) throws IOException {
        if (out.getPassword() != null) {
            return _encryptedWrap(out);
        }
        if (!out.isFile()) return (SheetOutput<OutputStream>) out;
        final OutputStream raw = Files.newOutputStream(((File) out.getRaw()).toPath());
        return out.isNamed() ? SheetOutput.target(raw, out.getName()) : SheetOutput.target(raw);
    }

    /**
     * Wraps an encrypted target — the writer streams plain OOXML into an
     * {@link EncryptedTempData} (AES-128-CBC with an in-memory random key, so
     * even if the temp survives a crash the bytes on disk are unreadable). On
     * close the temp is fed through {@link Encryptor} (agile, AES-256 + SHA-512)
     * and the encrypted bytes land on the original target.
     */
    private SheetOutput<OutputStream> _encryptedWrap(final SheetOutput<?> out) throws IOException {
        final EncryptedTempData tempData = new EncryptedTempData();
        final String password = out.getPassword();
        final EncryptionSpec spec = out.getEncryption() != null
                ? out.getEncryption() : EncryptionSpec.strong();
        // Captured locals — anonymous FilterOutputStream subclass shadows the
        // protected `out` field, so the SheetOutput must be referenced under a
        // distinct name to avoid the field shadowing.
        final SheetOutput<?> destination = out;
        final OutputStream tempStream;
        try {
            tempStream = new java.io.FilterOutputStream(tempData.getOutputStream()) {
                private boolean closed;
                @Override
                public void close() throws IOException {
                    if (closed) return;
                    closed = true;
                    super.close();
                    try {
                        _encryptToTarget(tempData, destination, password, spec);
                    } finally {
                        tempData.dispose();
                    }
                }
            };
        } catch (IOException | RuntimeException e) {
            tempData.dispose();
            throw e;
        }
        return out.isNamed()
                ? SheetOutput.target(tempStream, out.getName())
                : SheetOutput.target(tempStream);
    }

    @SuppressWarnings("unchecked")
    private void _encryptToTarget(final EncryptedTempData tempData, final SheetOutput<?> out,
                                  final String password, final EncryptionSpec spec) throws IOException {
        // SEC-18: for File targets, write encrypted bytes into a sibling temp file
        // and atomically rename onto the final target. Partial encrypted bytes
        // never appear at the target path; mid-write failures leave the original
        // target untouched. OutputStream targets cannot be made atomic by this
        // library — the caller owns the stream and must wrap accordingly.
        final File finalTarget = out.isFile() ? ((SheetOutput<File>) out).getRaw() : null;
        final File siblingTemp = finalTarget != null
                ? new File(finalTarget.getAbsoluteFile().getParentFile(),
                        finalTarget.getName() + ".enc-" + Long.toUnsignedString(System.nanoTime(), 36) + ".tmp")
                : null;
        final OutputStream target = (siblingTemp != null)
                ? new FileOutputStream(siblingTemp)
                : ((SheetOutput<OutputStream>) out).getRaw();
        boolean writeFinalized = false;
        IOException primary = null;
        try {
            try (InputStream plainIn = tempData.getInputStream();
                 POIFSFileSystem fs = new POIFSFileSystem();
                 OPCPackage opc = OPCPackage.open(plainIn)) {
                final EncryptionInfo info = spec.toEncryptionInfo();
                final Encryptor enc = Encryptor.getInstance(info);
                try {
                    enc.confirmPassword(password);
                    // Explicit close on the cipher stream is critical: OPCPackage.save
                    // only calls finish() on its ZipArchiveOutputStream wrap, never
                    // close(), so AgileCipherOutputStream never commits its
                    // EncryptionInfo / EncryptedPackage entries to the POIFS root.
                    try (OutputStream encStream = enc.getDataStream(fs)) {
                        opc.save(encStream);
                    }
                } catch (GeneralSecurityException e) {
                    throw new IOException("Failed to encrypt spreadsheet output", e);
                }
                fs.writeFilesystem(target);
                writeFinalized = true;
            } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
                throw new IOException("Failed to open temp data for encryption", e);
            }
        } catch (IOException e) {
            primary = e;
            throw e;
        } finally {
            // F3: addSuppressed pattern — preserve primary IOException
            if (siblingTemp != null) {
                try {
                    target.close();
                } catch (IOException closeEx) {
                    if (primary != null) primary.addSuppressed(closeEx);
                    else throw closeEx;
                }
                if (writeFinalized) {
                    try {
                        Files.move(siblingTemp.toPath(), finalTarget.toPath(),
                                java.nio.file.StandardCopyOption.ATOMIC_MOVE,
                                java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                    } catch (java.nio.file.AtomicMoveNotSupportedException ex) {
                        // Cross-filesystem fallback — non-atomic but the partial
                        // intermediate state is the sibling, not finalTarget.
                        Files.move(siblingTemp.toPath(), finalTarget.toPath(),
                                java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                    }
                } else {
                    // F4: mid-write failure — delete sibling, leave finalTarget untouched.
                    if (siblingTemp.exists() && !siblingTemp.delete() && log.isWarnEnabled()) {
                        log.warn("Failed to delete partial encrypted sibling: {}", siblingTemp);
                    }
                }
            }
        }
    }

    /**
     * Configurable features for spreadsheet read/write strategies.
     */
    public enum Feature implements FormatFeature {
        /**
         * Force POI User Model ({@code Sheet}/{@code Row}/{@code Cell}) for OOXML read/write,
         * bypassing SSML streaming.
         *
         * <p>POI User Model is also selected automatically when SSML is not applicable:
         * direct {@code Sheet} input/output, XLS files, non-OOXML {@code InputStream},
         * or non-XSSF custom workbooks. This flag is one of several triggers.
         *
         * <p>Default: disabled.
         */
        USE_POI_USER_MODEL(false),
        /**
         * Use file-backed shared string store to reduce heap usage for large files.
         * Applies to both read (SSMLSheetReader) and write (SSMLSheetWriter) paths.
         * Requires H2 at runtime.
         */
        FILE_BACKED_SHARED_STRINGS(false),
        /**
         * Encrypt file-backed shared string store at rest.
         * Effective only with {@link #FILE_BACKED_SHARED_STRINGS}.
         */
        ENCRYPT_FILE_BACKED_STORE(false),
        ;
        final boolean _defaultState;
        final int _mask;

        Feature(final boolean defaultState) {
            _defaultState = defaultState;
            _mask = 1 << ordinal();
        }

        public static int collectDefaults() {
            int flags = 0;
            for (Feature f : values()) {
                if (f.enabledByDefault()) {
                    flags |= f.getMask();
                }
            }
            return flags;
        }

        @Override
        public boolean enabledByDefault() {
            return _defaultState;
        }

        @Override
        public int getMask() {
            return _mask;
        }

        @Override
        public boolean enabledIn(final int flags) {
            return (flags & getMask()) != 0;
        }
    }
}
