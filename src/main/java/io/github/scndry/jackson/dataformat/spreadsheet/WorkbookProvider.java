package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;

public interface WorkbookProvider {

    Workbook create() throws IOException;
}
