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
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.ColumnNameResolver;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.generator.FormatVisitorWrapper;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.util.CellAddress;

import java.util.ArrayList;
import java.util.List;

@Slf4j
public final class SchemaGenerator {

    private final GeneratorSettings _generatorSettings;

    public SchemaGenerator() {
        _generatorSettings = new GeneratorSettings(CellAddress.A1, new StylesBuilder(), ColumnNameResolver.NULL);
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

    SpreadsheetSchema generate(final JavaType type, final DefaultSerializerProvider provider, final SerializerFactory factory)
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
        return new SpreadsheetSchema(columns, _generatorSettings._stylesBuilder, _generatorSettings._origin);
    }

    private void _verifyType(final JavaType type, final DefaultSerializerProvider provider) throws JsonMappingException {
        if (type.isArrayType() || type.isCollectionLikeType()) {
            throw _invalidSchemaDefinition(type, "can NOT be a Collection or array type");
        }
        if (!provider.getConfig().introspect(type).getClassAnnotations().has(DataGrid.class)) {
            throw _invalidSchemaDefinition(type, "MUST be annotated with `@" + DataGrid.class.getSimpleName() + '`');
        }
    }

    private JsonMappingException _invalidSchemaDefinition(final JavaType type, final String message) {
        return _invalidSchemaDefinition(type, "Root type of a schema " + message, null);
    }

    private JsonMappingException _invalidSchemaDefinition(final JavaType type, final Throwable cause) {
        return _invalidSchemaDefinition(type, cause.getMessage(), cause);
    }

    private JsonMappingException _invalidSchemaDefinition(final JavaType type, final String problem, final Throwable cause) {
        final String msg = String.format("Failed to generate schema of type '%s' for %s, problem: %s", SpreadsheetSchema.SCHEMA_TYPE,
                ClassUtil.getTypeDescription(type), problem);
        return InvalidDefinitionException.from((JsonGenerator) null, msg, type).withCause(cause);
    }

    @RequiredArgsConstructor
    static final class GeneratorSettings {

        private final CellAddress _origin;
        private final Styles.Builder _stylesBuilder;
        private final ColumnNameResolver _columnNameResolver;

        private GeneratorSettings with(final CellAddress origin) {
            return _origin.equals(origin) ? this : new GeneratorSettings(origin, _stylesBuilder, _columnNameResolver);
        }

        private GeneratorSettings with(final Styles.Builder styles) {
            return _stylesBuilder.equals(styles) ? this : new GeneratorSettings(_origin, styles, _columnNameResolver);
        }

        private GeneratorSettings with(final ColumnNameResolver resolver) {
            return _columnNameResolver.equals(resolver) ? this : new GeneratorSettings(_origin, _stylesBuilder, resolver);
        }
    }
}
