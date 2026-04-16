package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec;

/**
 * OOXML SpreadsheetML element names, attribute names, and namespace URIs
 * as defined in ECMA-376.
 * <p>
 * Single source of truth for all SpreadsheetML identifiers used in this library.
 */
public final class SpreadsheetML {

    private SpreadsheetML() {}

    // ---------------------------------------------------------------
    // Element local names
    // ---------------------------------------------------------------

    /** {@code <sheetData>} — sheet data container (§18.3.1.80) */
    public static final String SHEET_DATA = "sheetData";
    /** {@code <row>} — row element (§18.3.1.73) */
    public static final String ROW = "row";
    /** {@code <c>} — cell element (§18.3.1.4) */
    public static final String CELL = "c";
    /** {@code <v>} — cell value (§18.3.1.96) */
    public static final String VALUE = "v";
    /** {@code <f>} — cell formula (§18.3.1.40) */
    public static final String FORMULA = "f";
    /** {@code <is>} — inline string (§18.3.1.53) */
    public static final String INLINE_STRING = "is";
    /** {@code <t>} — text content (§18.4.12) */
    public static final String TEXT = "t";
    /** {@code <r>} — rich text run (§18.4.4) */
    public static final String RICH_TEXT_RUN = "r";
    /** {@code <sst>} — shared string table (§18.4.9) */
    public static final String SST = "sst";
    /** {@code <si>} — shared string item (§18.4.8) */
    public static final String STRING_ITEM = "si";
    /** {@code <workbook>} — workbook root (§18.2.27) */
    public static final String WORKBOOK = "workbook";
    /** {@code <workbookPr>} — workbook properties (§18.2.28) */
    public static final String WORKBOOK_PR = "workbookPr";
    /** {@code <sheet>} — sheet reference (§18.2.19) */
    public static final String SHEET = "sheet";
    /** {@code <sheets>} — sheet list container (§18.2.20) */
    public static final String SHEETS = "sheets";

    // ---------------------------------------------------------------
    // Attribute local names
    // ---------------------------------------------------------------

    /** {@code r} — cell reference on {@code <c>}, row index on {@code <row>} */
    public static final String ATTR_REF = "r";
    /** {@code t} — cell type on {@code <c>}, formula type on {@code <f>} */
    public static final String ATTR_TYPE = "t";
    /** {@code uniqueCount} — unique string count on {@code <sst>} */
    public static final String ATTR_UNIQUE_COUNT = "uniqueCount";
    /** {@code date1904} — date windowing on {@code <workbookPr>} */
    public static final String ATTR_DATE_1904 = "date1904";
    /** {@code name} — sheet name on {@code <sheet>} */
    public static final String ATTR_NAME = "name";
    /** {@code sheetId} — sheet ID on {@code <sheet>} */
    public static final String ATTR_SHEET_ID = "sheetId";

    // ---------------------------------------------------------------
    // Relationship namespace URIs
    // ---------------------------------------------------------------

    /** Transitional relationship namespace (OOXML 1st edition) */
    public static final String NS_REL_TRANSITIONAL =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /** Strict relationship namespace (ISO/IEC 29500 Strict) */
    public static final String NS_REL_STRICT =
            "http://purl.oclc.org/ooxml/officeDocument/relationships";
    /** Relationship attribute local name: {@code id} */
    public static final String ATTR_REL_ID = "id";
}
