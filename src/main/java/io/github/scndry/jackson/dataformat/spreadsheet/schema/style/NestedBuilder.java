package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

/**
 * Base class for nested builders that supports fluent chaining back to a parent builder via {@link #end()}.
 */
abstract class NestedBuilder<B extends Builder<?>> {

    private B _builder;

    protected final void setBuilder(final B builder) {
        _builder = builder;
    }

    public final B end() {
        return _builder;
    }
}
