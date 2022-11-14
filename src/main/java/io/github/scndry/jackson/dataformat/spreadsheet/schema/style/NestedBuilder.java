package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

abstract class NestedBuilder<B extends Builder<?>> {

    private B _builder;

    protected final void setBuilder(final B builder) {
        _builder = builder;
    }

    public final B end() {
        return _builder;
    }
}
