package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import java.util.function.Function;

/**
 * Base class for fluent enum property configurers that apply a selected value to a parent builder.
 */
abstract class PropertyConfigurer<E, B extends Builder<?>> {

    private final Function<E, B> _property;

    protected PropertyConfigurer(final Function<E, B> property) {
        _property = property;
    }

    protected final B apply(final E value) {
        return _property.apply(value);
    }
}
