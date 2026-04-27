package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.SerializationConfig;
import com.fasterxml.jackson.databind.exc.InvalidDefinitionException;
import com.fasterxml.jackson.databind.ser.DefaultSerializerProvider;
import com.fasterxml.jackson.databind.ser.SerializerFactory;
import com.fasterxml.jackson.databind.util.ClassUtil;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.feature.ConditionalFormattingConfigurer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.ColumnNameResolver;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.FormatVisitorWrapper;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.util.CellAddress;

import java.util.ArrayList;
import java.util.List;

@Slf4j
public final class SchemaGenerator {

    private final GeneratorSettings _generatorSettings;

    public SchemaGenerator() {
        _generatorSettings = new GeneratorSettings(
                CellAddress.A1,
                new StylesBuilder(),
                new ConditionalFormattingConfigurer(),
                ColumnNameResolver.NULL,
                SpreadsheetSchema.DEFAULT_FEATURES);
    }

    private SchemaGenerator(final GeneratorSettings generatorSettings) {
        _generatorSettings = generatorSettings;
    }

    public SchemaGenerator withOrigin(final CellAddress origin) {
        return new SchemaGenerator(_generatorSettings.with(origin));
    }

    public SchemaGenerator withStylesBuilder(final Styles.Builder builder) {
        return new SchemaGenerator(_generatorSettings.with(builder));
    }

    public SchemaGenerator withColumnNameResolver(final ColumnNameResolver resolver) {
        return new SchemaGenerator(_generatorSettings.with(resolver));
    }

    public SchemaGenerator withUseHeader(final boolean state) {
        return new SchemaGenerator(_generatorSettings.with(SpreadsheetSchema.FEATURE_USE_HEADER, state));
    }

    public SchemaGenerator withColumnReordering(final boolean state) {
        return new SchemaGenerator(_generatorSettings.with(SpreadsheetSchema.FEATURE_COLUMN_REORDERING, state));
    }

    public SchemaGenerator withConditionalFormattings(final ConditionalFormattingConfigurer builder) {
        return new SchemaGenerator(_generatorSettings.with(builder));
    }

    SpreadsheetSchema generate(
            final JavaType type,
            final DefaultSerializerProvider provider,
            final SerializerFactory factory)
            throws JsonMappingException {
        _verifyType(type, provider);
        final FormatVisitorWrapper visitor = new FormatVisitorWrapper();
        final SerializationConfig config = provider.getConfig()
                .withAttribute(ColumnNameResolver.class, _generatorSettings._columnNameResolver);
        final DefaultSerializerProvider instance = provider.createInstance(config, factory);
        try {
            instance.acceptJsonFormatVisitor(type, visitor);
        } catch (Exception e) {
            throw _invalidSchemaDefinition(type, e);
        }
        final List<Column> columns = new ArrayList<>();
        for (final Column column : visitor) {
            columns.add(column);
            if (log.isTraceEnabled()) {
                log.trace(column.toString());
            }
        }
        if (columns.isEmpty()) {
            throw _invalidSchemaDefinition(type,
                    "has no visible properties — check field visibility or add getters");
        }
        return new SpreadsheetSchema(
                columns,
                _generatorSettings._stylesBuilder,
                _generatorSettings._conditionalFormattings,
                _generatorSettings._origin,
                _generatorSettings._features);
    }

    private void _verifyType(
            final JavaType type,
            final DefaultSerializerProvider provider) throws JsonMappingException {
        if (type.isArrayType() || type.isCollectionLikeType()) {
            throw _invalidSchemaDefinition(type, "can NOT be a Collection or array type");
        }
        if (!provider.getConfig().introspect(type).getClassAnnotations().has(DataGrid.class)) {
            throw _invalidSchemaDefinition(
                    type,
                    "MUST be annotated with `@" + DataGrid.class.getSimpleName() + '`');
        }
    }

    private JsonMappingException _invalidSchemaDefinition(
            final JavaType type,
            final String message) {
        return _invalidSchemaDefinition(type, "Root type of a schema " + message, null);
    }

    private JsonMappingException _invalidSchemaDefinition(
            final JavaType type,
            final Throwable cause) {
        return _invalidSchemaDefinition(type, cause.getMessage(), cause);
    }

    private JsonMappingException _invalidSchemaDefinition(
            final JavaType type,
            final String problem,
            final Throwable cause) {
        final String msg = String.format(
                "Failed to generate schema of type '%s' for %s, problem: %s",
                SpreadsheetSchema.SCHEMA_TYPE,
                ClassUtil.getTypeDescription(type), problem);
        return InvalidDefinitionException.from((JsonGenerator) null, msg, type).withCause(cause);
    }

    static final class GeneratorSettings {

        private final CellAddress _origin;
        private final Styles.Builder _stylesBuilder;
        private final ConditionalFormattingConfigurer _conditionalFormattings;
        private final ColumnNameResolver _columnNameResolver;
        private final int _features;

        GeneratorSettings(final CellAddress origin, final Styles.Builder stylesBuilder,
                          final ConditionalFormattingConfigurer conditionalFormattings,
                          final ColumnNameResolver columnNameResolver, final int features) {
            _origin = origin;
            _stylesBuilder = stylesBuilder;
            _conditionalFormattings = conditionalFormattings;
            _columnNameResolver = columnNameResolver;
            _features = features;
        }

        private GeneratorSettings with(final CellAddress origin) {
            return _origin.equals(origin)
                    ? this
                    : new GeneratorSettings(origin, _stylesBuilder, _conditionalFormattings, _columnNameResolver, _features);
        }

        private GeneratorSettings with(final Styles.Builder styles) {
            return _stylesBuilder.equals(styles)
                    ? this
                    : new GeneratorSettings(_origin, styles, _conditionalFormattings, _columnNameResolver, _features);
        }

        private GeneratorSettings with(final ConditionalFormattingConfigurer conditionalFormattings) {
            return new GeneratorSettings(_origin, _stylesBuilder, conditionalFormattings, _columnNameResolver, _features);
        }

        private GeneratorSettings with(final ColumnNameResolver resolver) {
            return _columnNameResolver.equals(resolver)
                    ? this
                    : new GeneratorSettings(_origin, _stylesBuilder, _conditionalFormattings, resolver, _features);
        }

        private GeneratorSettings with(final int flag, final boolean state) {
            final int f = state ? (_features | flag) : (_features & ~flag);
            return f == _features
                    ? this
                    : new GeneratorSettings(_origin, _stylesBuilder, _conditionalFormattings, _columnNameResolver, f);
        }
    }
}
