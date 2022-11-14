package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.FormatSchema;
import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.PrettyPrinter;
import com.fasterxml.jackson.core.Version;
import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.ObjectWriter;
import com.fasterxml.jackson.databind.SequenceWriter;
import com.fasterxml.jackson.databind.SerializationConfig;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetGenerator;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;

public final class SpreadsheetWriter extends ObjectWriter {

    SpreadsheetWriter(final SpreadsheetMapper mapper, final SerializationConfig config,
                      final JavaType rootType, final PrettyPrinter pp) {
        super(mapper, config, rootType, pp);
    }

    SpreadsheetWriter(final SpreadsheetMapper mapper, final SerializationConfig config) {
        super(mapper, config);
    }

    SpreadsheetWriter(final SpreadsheetMapper mapper, final SerializationConfig config, final FormatSchema s) {
        super(mapper, config, s);
    }

    private SpreadsheetWriter(final SpreadsheetWriter base, final SerializationConfig config,
                              final GeneratorSettings genSettings, final Prefetch prefetch) {
        super(base, config, genSettings, prefetch);
    }

    private SpreadsheetWriter(final SpreadsheetWriter base, final SerializationConfig config) {
        super(base, config);
    }

    private SpreadsheetWriter(final SpreadsheetWriter base, final SpreadsheetFactory f) {
        super(base, f);
    }

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    /*
    /**********************************************************************
    /* Internal factory methods, for convenience
    /**********************************************************************
     */

    @Override
    protected ObjectWriter _new(final ObjectWriter base, final JsonFactory f) {
        return new SpreadsheetWriter((SpreadsheetWriter) base, (SpreadsheetFactory) f);
    }

    @Override
    protected ObjectWriter _new(final ObjectWriter base, final SerializationConfig config) {
        return new SpreadsheetWriter((SpreadsheetWriter) base, config);
    }

    @Override
    protected ObjectWriter _new(final GeneratorSettings genSettings, final Prefetch prefetch) {
        return new SpreadsheetWriter(this, _config, genSettings, prefetch);
    }

    /*
    /**********************************************************
    /* Factory methods for creating SheetGenerators
    /**********************************************************
     */

    public SheetGenerator createGenerator(final Sheet out) {
        _assertNotNull("out", out);
        return (SheetGenerator) _configureGenerator(generatorFactory().createGenerator(out));
    }

    public SheetGenerator createGenerator(final SheetOutput<?> out) throws IOException {
        _assertNotNull("out", out);
        return (SheetGenerator) _configureGenerator(generatorFactory().createGenerator(out));
    }

    /*
    /**********************************************************
    /* Factory methods for sequence writers
    /**********************************************************
     */

    public SequenceWriter writeValues(final Sheet out) throws IOException {
        return _newSequenceWriter(false, createGenerator(out), true);
    }

    public SequenceWriter writeValues(final SheetOutput<?> out) throws IOException {
        return _newSequenceWriter(false, createGenerator(out), true);
    }

    public SequenceWriter writeValuesAsArray(final Sheet out) throws IOException {
        return _newSequenceWriter(true, createGenerator(out), true);
    }

    public SequenceWriter writeValuesAsArray(final SheetOutput<?> out) throws IOException {
        return _newSequenceWriter(true, createGenerator(out), true);
    }

    /*
    /**********************************************************
    /* Serialization methods, others
    /**********************************************************
     */

    public void writeValue(final Sheet out, final Object value) throws IOException {
        _writeValueAndClose(createGenerator(out), value);
    }

    public void writeValue(final SheetOutput<?> out, final Object value) throws IOException {
        _writeValueAndClose(createGenerator(out), value);
    }

    /*
    /**********************************************************
    /* Other public methods
    /**********************************************************
     */

    public SpreadsheetFactory generatorFactory() {
        return (SpreadsheetFactory) _generatorFactory;
    }
}
