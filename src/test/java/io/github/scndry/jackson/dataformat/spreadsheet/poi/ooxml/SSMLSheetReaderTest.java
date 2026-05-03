package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.File;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.PoiVersionProbe;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReaderTestBase;

class SSMLSheetReaderTest extends SheetReaderTestBase {

    @Test
    void transitional() throws Exception {
        File file = createHeadlessFixture();
        SSMLWorkbook workbook = SSMLWorkbook.create(file);
        PackagePart part = workbook.getWorksheetPartAt(0);
        reader = new SSMLSheetReader(part, workbook, false, false);
        testSheetReader();
    }

    @Test
    void strict() throws Exception {
        // Strict OOXML reading requires POI 5.1.0+ (see GUIDE.md, POICompat.STRICT_FORMAT_METHOD).
        // POI 4.x cannot resolve the worksheet relationship in a strict-format package.
        Assumptions.assumeTrue(PoiVersionProbe.isPoi510OrLater(),
                "Strict OOXML reading requires POI 5.1.0+");

        SSMLWorkbook workbook = SSMLWorkbook.create(strictSource);
        PackagePart part = workbook.getWorksheetPartAt(0);
        reader = new SSMLSheetReader(part, workbook, false, false);
        testSheetReader();
    }
}
