package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.databind.*;
import com.fasterxml.jackson.databind.cfg.MapperBuilder;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.ColumnNameResolver;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetGenerator;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

@SuppressWarnings("java:S2177")
public final class SpreadsheetMapper extends ObjectMapper {

    private transient SchemaGenerator _schemaGenerator;

    public SpreadsheetMapper() {
        this(new SpreadsheetFactory());
    }

    public SpreadsheetMapper(final SpreadsheetFactory f) {
        super(f);
        registerModule(new SpreadsheetModule());
        _schemaGenerator = new SchemaGenerator();
    }

    private SpreadsheetMapper(final SpreadsheetMapper src) {
        super(src);
        _schemaGenerator = src._schemaGenerator;
    }

    public static Builder builder() {
        return new Builder(new SpreadsheetMapper());
    }

    public static Builder builder(final SpreadsheetFactory streamFactory) {
        return new Builder(new SpreadsheetMapper(streamFactory));
    }

    public Builder rebuild() {
        return new Builder(copy());
    }

    @Override
    public SpreadsheetMapper copy() {
        _checkInvalidCopy(SpreadsheetMapper.class);
        return new SpreadsheetMapper(this);
    }

    @Override
    public SpreadsheetMapper copyWith(final JsonFactory factory) {
        return (SpreadsheetMapper) super.copyWith(factory);
    }

    /*
    /**********************************************************
    /* SpreadsheetReader/SpreadsheetWriter implementations
    /**********************************************************
     */

    @Override
    protected SpreadsheetReader _newReader(DeserializationConfig config) {
        return new SpreadsheetReader(this, config);
    }

    @Override
    protected SpreadsheetReader _newReader(final DeserializationConfig config, final JavaType valueType,
                                           final Object valueToUpdate, final FormatSchema schema,
                                           final InjectableValues injectableValues) {
        return new SpreadsheetReader(this, config, valueType, valueToUpdate, schema, injectableValues);
    }

    @Override
    protected SpreadsheetWriter _newWriter(final SerializationConfig config) {
        return new SpreadsheetWriter(this, config);
    }

    @Override
    protected SpreadsheetWriter _newWriter(final SerializationConfig config, final FormatSchema schema) {
        return new SpreadsheetWriter(this, config, schema);
    }

    @Override
    protected SpreadsheetWriter _newWriter(final SerializationConfig config, final JavaType rootType,
                                           final PrettyPrinter pp) {
        return new SpreadsheetWriter(this, config, rootType, pp);
    }

    /*
    /**********************************************************
    /* Versioned impl
    /**********************************************************
     */

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    /*
    /**********************************************************
    /* Module registration, discovery
    /**********************************************************
     */

    @Override
    public SpreadsheetMapper registerModule(final com.fasterxml.jackson.databind.Module module) {
        return (SpreadsheetMapper) super.registerModule(module);
    }

    @Override
    public SpreadsheetMapper registerModules(final com.fasterxml.jackson.databind.Module... modules) {
        return (SpreadsheetMapper) super.registerModules(modules);
    }

    @Override
    public SpreadsheetMapper registerModules(final Iterable<? extends com.fasterxml.jackson.databind.Module> modules) {
        return (SpreadsheetMapper) super.registerModules(modules);
    }

    /*
    /**********************************************************
    /* Factory methods for creating SheetGenerators
    /**********************************************************
     */

    public SheetGenerator createGenerator(final Sheet out) {
        _assertNotNull("out", out);
        final SheetGenerator g = tokenStreamFactory().createGenerator(out);
        _serializationConfig.initialize(g);
        return g;
    }

    public SheetGenerator createGenerator(final SheetOutput<?> out) throws IOException {
        _assertNotNull("out", out);
        final SheetGenerator g = tokenStreamFactory().createGenerator(out);
        _serializationConfig.initialize(g);
        return g;
    }

    @Override
    public SheetGenerator createGenerator(final OutputStream out) throws IOException {
        return (SheetGenerator) super.createGenerator(out);
    }

    @Override
    public SheetGenerator createGenerator(final OutputStream out, final JsonEncoding enc) throws IOException {
        return (SheetGenerator) super.createGenerator(out, enc);
    }

    @Override
    public SheetGenerator createGenerator(final File outputFile, final JsonEncoding enc) throws IOException {
        return (SheetGenerator) super.createGenerator(outputFile, enc);
    }

    /*
    /**********************************************************
    /* Factory methods for creating SheetParsers
    /**********************************************************
     */

    public SheetParser createParser(final Sheet src) {
        _assertNotNull("src", src);
        return (SheetParser) _deserializationConfig.initialize(tokenStreamFactory().createParser(src));
    }

    public SheetParser createParser(final SheetInput<?> src) throws IOException {
        _assertNotNull("src", src);
        return (SheetParser) _deserializationConfig.initialize(tokenStreamFactory().createParser(src));
    }

    @Override
    public SheetParser createParser(final File src) throws IOException {
        return (SheetParser) super.createParser(src);
    }

    @Override
    public SheetParser createParser(final InputStream in) throws IOException {
        return (SheetParser) super.createParser(in);
    }

    /*
    /**********************************************************
    /* Configuration, simple features: SheetParser.Feature
    /**********************************************************
     */

    public SpreadsheetMapper configure(final SheetParser.Feature f, final boolean state) {
        tokenStreamFactory().configure(f, state);
        return this;
    }

    public SpreadsheetMapper enable(final SheetParser.Feature... features) {
        for (SheetParser.Feature f : features) {
            tokenStreamFactory().enable(f);
        }
        return this;
    }

    public SpreadsheetMapper disable(final SheetParser.Feature... features) {
        for (SheetParser.Feature f : features) {
            tokenStreamFactory().disable(f);
        }
        return this;
    }

    /*
    /**********************************************************
    /* Configuration, schema generation
    /**********************************************************
     */

    public SchemaGenerator getSchemaGenerator() {
        return _schemaGenerator;
    }

    public SpreadsheetMapper setSchemaGenerator(final SchemaGenerator generator) {
        _assertNotNull("generator", generator);
        _schemaGenerator = generator;
        return this;
    }

    public SpreadsheetMapper setOrigin(final int row, final int column) {
        return setOrigin(new CellAddress(row, column));
    }

    public SpreadsheetMapper setOrigin(final String address) {
        _assertNotNull("address", address);
        return setSchemaGenerator(_schemaGenerator.withOrigin(new CellAddress(address)));
    }

    public SpreadsheetMapper setOrigin(final CellAddress address) {
        _assertNotNull("address", address);
        return setSchemaGenerator(_schemaGenerator.withOrigin(address));
    }

    public SpreadsheetMapper setStylesBuilder(final Styles.Builder builder) {
        _assertNotNull("builder", builder);
        return setSchemaGenerator(_schemaGenerator.withStylesBuilder(builder));
    }

    public SpreadsheetMapper setColumnNameResolver(final ColumnNameResolver resolver) {
        _assertNotNull("resolver", resolver);
        return setSchemaGenerator(_schemaGenerator.withColumnNameResolver(resolver));
    }

    /*
    /**********************************************************
    /* Configuration, other
    /**********************************************************
     */

    @Override
    public SpreadsheetFactory tokenStreamFactory() {
        return (SpreadsheetFactory) super.tokenStreamFactory();
    }

    /*
    /**********************************************************
    /* Public API, deserialization,
    /* convenience methods
    /**********************************************************
     */

    public <T> T readValue(final Sheet src, final Class<T> valueType) throws IOException {
        return sheetReaderFor(valueType).readValue(src);
    }

    public <T> T readValue(final SheetInput<?> src, final Class<T> valueType) throws IOException {
        return sheetReaderFor(valueType).readValue(src);
    }

    @Override
    public <T> T readValue(final File src, final Class<T> valueType) throws IOException {
        return sheetReaderFor(valueType).readValue(src);
    }

    @Override
    public <T> T readValue(final InputStream src, final Class<T> valueType) throws IOException {
        return sheetReaderFor(valueType).readValue(src);
    }

    public <T> List<T> readValues(final Sheet src, final Class<T> valueType) throws IOException {
        try (MappingIterator<T> iterator = sheetReaderFor(valueType).readValues(src)) {
            return iterator.readAll();
        }
    }

    public <T> List<T> readValues(final SheetInput<?> src, final Class<T> valueType) throws IOException {
        try (MappingIterator<T> iterator = sheetReaderFor(valueType).readValues(src)) {
            return iterator.readAll();
        }
    }

    public <T> List<T> readValues(final File src, final Class<T> valueType) throws IOException {
        try (MappingIterator<T> iterator = sheetReaderFor(valueType).readValues(src)) {
            return iterator.readAll();
        }
    }

    public <T> List<T> readValues(final InputStream src, final Class<T> valueType) throws IOException {
        try (MappingIterator<T> iterator = sheetReaderFor(valueType).readValues(src)) {
            return iterator.readAll();
        }
    }

    /*
    /**********************************************************
    /* Public API: serialization
    /* convenience methods
    /**********************************************************
     */

    public void writeValue(final Sheet out, final Object value) throws IOException {
        _verifyValueType(value);
        writeValue(out, value, value.getClass());
    }

    public void writeValue(final SheetOutput<?> out, final Object value) throws IOException {
        _verifyValueType(value);
        writeValue(out, value, value.getClass());
    }

    @Override
    public void writeValue(final File out, final Object value) throws IOException {
        _verifyValueType(value);
        writeValue(out, value, value.getClass());
    }

    @Override
    public void writeValue(final OutputStream out, final Object value) throws IOException {
        _verifyValueType(value);
        writeValue(out, value, value.getClass());
    }

    public void writeValue(final Sheet out, final Object value, final Class<?> valueType) throws IOException {
        sheetWriterFor(valueType).writeValue(out, value);
    }

    public void writeValue(final SheetOutput<?> out, final Object value, final Class<?> valueType) throws IOException {
        sheetWriterFor(valueType).writeValue(out, value);
    }

    public void writeValue(final File out, final Object value, final Class<?> valueType) throws IOException {
        sheetWriterFor(valueType).writeValue(out, value);
    }

    public void writeValue(final OutputStream out, final Object value, final Class<?> valueType) throws IOException {
        sheetWriterFor(valueType).writeValue(out, value);
    }

    @Override
    public byte[] writeValueAsBytes(final Object value) throws JsonProcessingException {
        return writeValueAsBytes(value, value.getClass());
    }

    public byte[] writeValueAsBytes(final Object value, final Class<?> valueType) throws JsonProcessingException {
        return sheetWriterFor(valueType).writeValueAsBytes(value);
    }

    /*
    /**********************************************************
    /* Public API: constructing SpreadsheetWriters
    /* for more advanced configuration
    /**********************************************************
     */

    @Override
    public SpreadsheetWriter writer(final FormatSchema schema) {
        return (SpreadsheetWriter) super.writer(schema);
    }

    /**
     * Convenience method, equivalent in function to:
     * <pre>{@code
     * writer(sheetSchemaFor(type));
     * }</pre>
     */
    public SpreadsheetWriter sheetWriterFor(final Class<?> type) throws JsonMappingException {
        return writer(sheetSchemaFor(type));
    }

    /*
    /**********************************************************
    /* Public API: constructing SpreadsheetReaders
    /* for more advanced configuration
    /**********************************************************
     */

    @Override
    public SpreadsheetReader reader(final FormatSchema schema) {
        return (SpreadsheetReader) super.reader(schema);
    }

    /**
     * Convenience method, equivalent in function to:
     * <pre>{@code
     * reader(sheetSchemaFor(type)).forType(type);
     * }</pre>
     */
    public SpreadsheetReader sheetReaderFor(final Class<?> type) throws JsonMappingException {
        return reader(sheetSchemaFor(type)).forType(type);
    }

    /*
    /**********************************************************
    /* Public API: SpreadsheetSchema generation
    /**********************************************************
     */

    public SpreadsheetSchema sheetSchemaFor(final Class<?> type) throws JsonMappingException {
        return _schemaGenerator.generate(constructType(type), _serializerProvider(_serializationConfig), _serializerFactory);
    }

    /*
    /**********************************************************
    /* Internal methods, other
    /**********************************************************
     */

    private void _verifyValueType(final Object value) {
        // Type can NOT be a Collection or array type
        final JavaType type = constructType(value.getClass());
        if (type.isArrayType() || type.isCollectionLikeType()) {
            throw new IllegalArgumentException("`valueType` MUST be specified to write a value of a Collection or array type");
        }
    }

    public static final class Builder extends MapperBuilder<SpreadsheetMapper, Builder> {

        private Builder(final SpreadsheetMapper mapper) {
            super(mapper);
        }

        public Builder enable(final SheetParser.Feature... features) {
            _mapper.enable(features);
            return _this();
        }

        public Builder disable(final SheetParser.Feature... features) {
            _mapper.disable(features);
            return _this();
        }

        public Builder configure(final SheetParser.Feature feature, final boolean state) {
            _mapper.configure(feature, state);
            return _this();
        }

        public Builder schemaGenerator(final SchemaGenerator generator) {
            _mapper.setSchemaGenerator(generator);
            return _this();
        }

        public Builder origin(final int row, final int column) {
            return origin(new CellAddress(row, column));
        }

        public Builder origin(final String address) {
            return origin(new CellAddress(address));
        }

        public Builder origin(final CellAddress address) {
            _mapper.setOrigin(address);
            return _this();
        }

        public Builder stylesBuilder(final Styles.Builder builder) {
            _mapper.setStylesBuilder(builder);
            return _this();
        }

        public Builder columnNameResolver(final ColumnNameResolver resolver) {
            _mapper.setColumnNameResolver(resolver);
            return _this();
        }
    }
}
