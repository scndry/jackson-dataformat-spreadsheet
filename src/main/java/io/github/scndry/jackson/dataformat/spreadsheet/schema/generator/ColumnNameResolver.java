package io.github.scndry.jackson.dataformat.spreadsheet.schema.generator;

import java.lang.annotation.Annotation;
import java.util.function.Function;

import com.fasterxml.jackson.databind.BeanProperty;

/**
 * Strategy interface for resolving spreadsheet column names from bean properties.
 */
public interface ColumnNameResolver {

    ColumnNameResolver NULL = prop -> null;

    String resolve(BeanProperty prop);

    static <A extends Annotation, T>
    AnnotatedNameResolver<A> annotated(final Class<A> type, final Function<T, String> nameMapper) {
        return AnnotatedNameResolver.forValue(type, nameMapper);
    }
}
