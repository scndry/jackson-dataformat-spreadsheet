package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import org.apache.poi.ss.usermodel.Workbook;

interface Builder<O> {

    O build(Workbook workbook);

    Builder<O> copy();
}
