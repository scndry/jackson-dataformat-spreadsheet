package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xssf.usermodel.XSSFRelation;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;

/**
 * Lightweight OOXML workbook navigator. Parses {@code workbook.xml} via {@link XmlElementReader}
 * to extract the sheet list and date1904 setting without loading the full workbook object model.
 *
 * @see XSSFCorePart
 * @see XmlElementReader
 */
public final class SSMLWorkbook implements AutoCloseable {

    private static final Matcher START_WORKBOOK_PR =
            Matcher.startElement(SpreadsheetML.WORKBOOK_PR);
    private static final Matcher START_SHEET = Matcher.startElement(SpreadsheetML.SHEET);
    private static final Matcher END_SHEETS = Matcher.endElement(SpreadsheetML.SHEETS);

    private final XSSFCorePart _corePart;
    private final boolean _date1904;
    private final PackageRelationshipCollection _worksheetRels;
    private final List<SheetEntry> _sheets;

    private SSMLWorkbook(final XSSFCorePart corePart) {
        _corePart = corePart;
        _sheets = new ArrayList<>();
        boolean date1904 = false;
        try (XmlElementReader reader = new XmlElementReader(_corePart.getInputStream())) {
            reader.navigateTo(SpreadsheetML.WORKBOOK);
            final Matcher hit = reader.nextUntil(START_WORKBOOK_PR, START_SHEET, END_SHEETS);
            if (hit == START_WORKBOOK_PR) {
                final String val = reader.attribute(SpreadsheetML.ATTR_DATE_1904);
                date1904 = "1".equals(val) || "true".equals(val);
                final Matcher next = reader.nextUntil(START_SHEET, END_SHEETS);
                if (next == START_SHEET) {
                    _addSheet(reader);
                }
            } else if (hit == START_SHEET) {
                _addSheet(reader);
            }
            while (true) {
                final Matcher m = reader.nextUntil(START_SHEET, END_SHEETS);
                if (m == null || m.isEndElement()) break;
                _addSheet(reader);
            }
        }
        _date1904 = date1904;
        _worksheetRels = _corePart.getRelationships(XSSFRelation.WORKSHEET);
    }

    private void _addSheet(final XmlElementReader reader) {
        String rId = reader.attribute(SpreadsheetML.NS_REL_TRANSITIONAL, SpreadsheetML.ATTR_REL_ID);
        if (rId == null) {
            rId = reader.attribute(SpreadsheetML.NS_REL_STRICT, SpreadsheetML.ATTR_REL_ID);
        }
        _sheets.add(new SheetEntry(
                reader.attribute(SpreadsheetML.ATTR_NAME),
                reader.attribute(SpreadsheetML.ATTR_SHEET_ID),
                rId));
    }

    public static SSMLWorkbook create(final File source) {
        return _create(PackageUtil.extractCorePart(source));
    }

    public static SSMLWorkbook create(final InputStream source) {
        return _create(PackageUtil.extractCorePart(source));
    }

    private static SSMLWorkbook _create(final PackagePart corePart) {
        return new SSMLWorkbook(new XSSFCorePart(corePart));
    }

    public boolean isDate1904() {
        return _date1904;
    }

    public PackagePart getSharedStringsPart() {
        return _corePart.getRelatedPart(XSSFRelation.SHARED_STRINGS);
    }

    public PackagePart getStylesPart() {
        return _corePart.getRelatedPart(XSSFRelation.STYLES);
    }

    // Index-based: throws on out of range (programming error, like List.get)
    public PackagePart getWorksheetPartAt(final int index) {
        _validateSheetIndex(index);
        return _worksheetPart(_sheets.get(index));
    }

    // Name-based: returns null if not found (lookup failure, like Map.get)
    public PackagePart getWorksheetPart(final String name) {
        for (final SheetEntry sheet : _sheets) {
            if (sheet._name.equalsIgnoreCase(name)) {
                return _worksheetPart(sheet);
            }
        }
        return null;
    }

    @Override
    public void close() throws IOException {
        _corePart.getPackage().close();
    }

    private PackagePart _worksheetPart(final SheetEntry sheet) {
        final String id = Optional.ofNullable(sheet._rId).orElseGet(() -> "rId" + sheet._sheetId);
        final PackageRelationship rel = _worksheetRels.getRelationshipByID(id);
        return _corePart.getRelatedPart(rel);
    }

    private void _validateSheetIndex(final int index) {
        final int lastSheetIx = _sheets.size() - 1;
        if (index < 0 || index > lastSheetIx) {
            String range = "(0.." + lastSheetIx + ")";
            if (lastSheetIx == -1) {
                range = "(no sheets)";
            }
            throw new IllegalArgumentException("Sheet index (" + index + ") is out of range " +
                    range);
        }
    }

    private static final class SheetEntry {
        final String _name;
        final String _sheetId;
        final String _rId;

        SheetEntry(final String name, final String sheetId, final String rId) {
            _name = name;
            _sheetId = sheetId;
            _rId = rId;
        }
    }
}
