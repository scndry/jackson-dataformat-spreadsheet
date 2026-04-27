package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.filesystem.FileMagic;
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
        final SheetInput<?> source = _preferRawAsFile(src);
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
        if (Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags)) {
            return _createPOISheetReader(WorkbookFactory.create(src.getRaw()), src);
        }
        if (PackageUtil.isOOXML(src.getRaw())) {
            return _createSSMLSheetReader(SSMLWorkbook.create(src.getRaw()), src);
        }
        return _createPOISheetReader(WorkbookFactory.create(src.getRaw()), src);
    }

    private SheetReader _createInputStreamSheetReader(
            final SheetInput<InputStream> src) throws IOException {
        if (Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags)) {
            return _createPOISheetReader(WorkbookFactory.create(src.getRaw()), src);
        }
        if (PackageUtil.isOOXML(src.getRaw())) {
            return _createSSMLSheetReader(SSMLWorkbook.create(src.getRaw()), src);
        }
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

    @SuppressWarnings("unchecked")
    private SheetInput<?> _preferRawAsFile(final SheetInput<?> src) throws IOException {
        if (src.isFile()) return src;
        if (Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags)) return src;
        final InputStream raw = FileMagic.prepareToCheckMagic(
                ((SheetInput<InputStream>) src).getRaw());
        if (!PackageUtil.isOOXML(raw)) {
            return src.isNamed()
                    ? SheetInput.source(raw, src.getName())
                    : SheetInput.source(raw, src.getIndex());
        }
        final File file = POICompat.createSecureTempFile("jackson-spreadsheet-input-", ".xlsx").toFile();
        try (FileOutputStream out = new FileOutputStream(file)) {
            final byte[] buf = new byte[8192];
            int n;
            while ((n = raw.read(buf)) != -1) {
                out.write(buf, 0, n);
            }
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
        if (Feature.USE_POI_USER_MODEL.enabledIn(_featureFlags)) {
            return _createPOISheetWriter(sheet, out.getRaw());
        }
        if (sheet instanceof XSSFSheet) {
            return _createSSMLSheetWriter(out.getRaw(), sheet);
        }
        if (log.isDebugEnabled()) {
            log.debug("Sheet is not XSSFSheet ({}); falling back to POISheetWriter",
                    sheet.getClass().getSimpleName());
        }
        return _createPOISheetWriter(sheet, out.getRaw());
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

    @SuppressWarnings("unchecked")
    private SheetOutput<OutputStream> _rawAsOutputStream(
            final SheetOutput<?> out) throws IOException {
        if (!out.isFile()) return (SheetOutput<OutputStream>) out;
        final OutputStream raw = Files.newOutputStream(((File) out.getRaw()).toPath());
        return out.isNamed() ? SheetOutput.target(raw, out.getName()) : SheetOutput.target(raw);
    }

    /**
     * Configurable features for spreadsheet read/write strategies.
     */
    public enum Feature implements FormatFeature {
        /**
         * Use POI's User Model ({@code Sheet}/{@code Row}/{@code Cell}) for all read/write operations,
         * bypassing SSML streaming. Default: disabled (SSML is used for OOXML).
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
