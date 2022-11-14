package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonToken;
import com.fasterxml.jackson.databind.BeanDescription;
import com.fasterxml.jackson.databind.DeserializationConfig;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;
import com.fasterxml.jackson.databind.deser.BeanDeserializerModifier;
import com.fasterxml.jackson.databind.deser.std.DelegatingDeserializer;
import com.fasterxml.jackson.databind.util.Annotations;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import java.io.IOException;

public final class DataGridBeanDeserializer extends DelegatingDeserializer {

    public DataGridBeanDeserializer(final JsonDeserializer<?> d) {
        super(d);
    }

    @Override
    public Object deserialize(final JsonParser p, final DeserializationContext ctxt) throws IOException {
        if (p.currentToken() == JsonToken.VALUE_NULL) {
            return null;
        }
        return super.deserialize(p, ctxt);
    }

    @Override
    protected JsonDeserializer<?> newDelegatingInstance(final JsonDeserializer<?> newDelegatee) {
        throw new IllegalStateException("Internal error: should never get called");
    }

    public static final class Modifier extends BeanDeserializerModifier {
        @Override
        public JsonDeserializer<?> modifyDeserializer(final DeserializationConfig config, final BeanDescription beanDesc, final JsonDeserializer<?> deserializer) {
            final Annotations annotations = beanDesc.getClassAnnotations();
            if (annotations.has(DataGrid.class)) {
                return new DataGridBeanDeserializer(deserializer);
            }
            return deserializer;
        }
    }
}
