package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Base interface for style builders that produce POI objects from a {@link Workbook}.
 */
interface Builder<O> {

    O build(Workbook workbook);

    Builder<O> copy();
}
