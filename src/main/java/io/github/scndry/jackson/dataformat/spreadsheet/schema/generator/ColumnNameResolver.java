package io.github.scndry.jackson.dataformat.spreadsheet.schema.generator;

import com.fasterxml.jackson.databind.BeanProperty;

public interface ColumnNameResolver {

    ColumnNameResolver NULL = prop -> null;

    String resolve(BeanProperty prop);
}
