package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public final class SSMLWorkbook implements AutoCloseable {

    private final XSSFCorePart _corePart;
    private final boolean _date1904;
    private final PackageRelationshipCollection _worksheetRels;
    private final List<CTSheet> _sheets;

    private SSMLWorkbook(final XSSFCorePart corePart) {
        _corePart = corePart;
        final WorkbookDocument doc;
        try (XmlElementReader reader = new XmlElementReader(_corePart.getInputStream())) {
            doc = reader.collect();
        }
        final CTWorkbook workbook = doc.getWorkbook();
        _date1904 = workbook.getWorkbookPr().getDate1904();
        _worksheetRels = _corePart.getRelationshipsByType(XSSFRelation.WORKSHEET);
        _sheets = workbook.getSheets().getSheetList();
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

    public PackagePart getSharedStrings() {
        return _corePart.getRelatedPart(XSSFRelation.SHARED_STRINGS);
    }

    public PackagePart getStyles() {
        return _corePart.getRelatedPart(XSSFRelation.STYLES);
    }

    public PackagePart getSheetAt(final int index) {
        _validateSheetIndex(index);
        final CTSheet sheet = _sheets.get(index);
        final PackageRelationship rel = _worksheetRels.getRelationshipByID(sheet.getId());
        return _corePart.getRelatedPart(rel);
    }

    public PackagePart getSheet(final String name) {
        for (final CTSheet sheet : _sheets) {
            if (sheet.getName().equalsIgnoreCase(name)) {
                final PackageRelationship rel = _worksheetRels.getRelationshipByID(sheet.getId());
                return _corePart.getRelatedPart(rel);
            }
        }
        return null;
    }

    @Override
    public void close() throws IOException {
        _corePart.getPackage().close();
    }

    private void _validateSheetIndex(final int index) {
        final int lastSheetIx = _sheets.size() - 1;
        if (index < 0 || index > lastSheetIx) {
            String range = "(0.." + lastSheetIx + ")";
            if (lastSheetIx == -1) {
                range = "(no sheets)";
            }
            throw new IllegalArgumentException("Sheet index (" + index + ") is out of range " + range);
        }
    }
}
