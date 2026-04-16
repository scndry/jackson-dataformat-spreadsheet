package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Functional interface for creating Apache POI {@link Workbook}
 * instances. Used by {@link SpreadsheetFactory} to control which
 * workbook implementation (e.g. SXSSF, XSSF, HSSF) is used for
 * writing.
 *
 * @see SpreadsheetFactory
 */
public interface WorkbookProvider {

    Workbook create() throws IOException;
}
