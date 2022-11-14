package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.io.IOContext;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.PackageUtil;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.SSMLSheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.SSMLWorkbook;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ss.POISheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ss.POISheetWriter;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetGenerator;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetWriter;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

@SuppressWarnings("java:S2177")
public final class SpreadsheetFactory extends JsonFactory {

    public static final String FORMAT_NAME = "spreadsheet";
    public static final int DEFAULT_SHEET_PARSER_FEATURE_FLAGS = SheetParser.Feature.collectDefaults();

    private final transient WorkbookProvider _workbookProvider;
    private int _sheetParserFeatures;

    public SpreadsheetFactory() {
        this(SXSSFWorkbook::new, DEFAULT_SHEET_PARSER_FEATURE_FLAGS);
    }

    public SpreadsheetFactory(final WorkbookProvider workbookProvider, final int sheetParserFeatures) {
        _workbookProvider = workbookProvider;
        _sheetParserFeatures = sheetParserFeatures;
    }

    public SpreadsheetFactory(final SpreadsheetFactory base) {
        _workbookProvider = base._workbookProvider;
        _sheetParserFeatures = base._sheetParserFeatures;
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
    /* Parser factories
    /**********************************************************
     */

    public SheetParser createParser(final Sheet src) {
        final IOContext ctxt = _createContext(_createContentReference(src), false);
        return _createParser(new POISheetReader(src), ctxt);
    }

    @SuppressWarnings("unchecked")
    public SheetParser createParser(final SheetInput<?> src) throws IOException {
        final IOContext ctxt = _createContext(_createContentReference(src), src.isFile());
        final SheetReader reader;
        if (src.isFile()) {
            reader = _createFileSheetReader((SheetInput<File>) src);
        } else {
            reader = _createInputStreamSheetReader((SheetInput<InputStream>) src);
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
        final IOContext ctxt = _createContext(_createContentReference(out), out.isFile());
        return _createGenerator(_createSheetWriter(out), ctxt);
    }

    @Override
    public SheetGenerator createGenerator(final File out, final JsonEncoding enc) throws IOException {
        return createGenerator(SheetOutput.target(out));
    }

    @Override
    public SheetGenerator createGenerator(final OutputStream out, final JsonEncoding enc) throws IOException {
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
        if (PackageUtil.isOOXML(src.getRaw())) {
            return _createSSMLSheetReader(SSMLWorkbook.create(src.getRaw()), src);
        }
        return _createPOISheetReader(WorkbookFactory.create(src.getRaw()), src);
    }

    private SheetReader _createInputStreamSheetReader(final SheetInput<InputStream> src) throws IOException {
        if (PackageUtil.isOOXML(src.getRaw())) {
            return _createSSMLSheetReader(SSMLWorkbook.create(src.getRaw()), src);
        }
        return _createPOISheetReader(WorkbookFactory.create(src.getRaw()), src);
    }

    private SSMLSheetReader _createSSMLSheetReader(final SSMLWorkbook workbook, final SheetInput<?> src) {
        final PackagePart worksheetPart = src.isNamed() ? workbook.getSheet(src.getName()) : workbook.getSheetAt(src.getIndex());
        if (worksheetPart == null) {
            throw new IllegalArgumentException("No sheet for " + src);
        }
        return new SSMLSheetReader(worksheetPart, workbook);
    }

    private POISheetReader _createPOISheetReader(final Workbook workbook, final SheetInput<?> src) {
        final Sheet sheet = src.isNamed() ? workbook.getSheet(src.getName()) : workbook.getSheetAt(src.getIndex());
        if (sheet == null) {
            throw new IllegalArgumentException("No sheet for " + src);
        }
        return new POISheetReader(sheet);
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
    private SheetWriter _createSheetWriter(final SheetOutput<?> out) throws IOException {
        final Workbook workbook = _workbookProvider.create();
        final Sheet sheet;
        if (out.isNamed()) {
            sheet = workbook.createSheet(out.getName());
        } else {
            sheet = workbook.createSheet();
        }
        return new POISheetWriter(sheet);
    }
}
