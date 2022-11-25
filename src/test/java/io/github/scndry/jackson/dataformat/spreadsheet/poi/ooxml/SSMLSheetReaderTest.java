package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReaderTestBase;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.junit.jupiter.api.Test;

class SSMLSheetReaderTest extends SheetReaderTestBase {

    @Test
    void testTransitional() throws Exception {
        final SSMLWorkbook workbook = SSMLWorkbook.create(transitionalSource);
        final PackagePart worksheetPart = workbook.getWorksheetPartAt(0);
        reader = new SSMLSheetReader(worksheetPart, workbook);
        testSheetReader();
    }

    @Test
    void testStrict() throws Exception {
        final SSMLWorkbook workbook = SSMLWorkbook.create(strictSource);
        final PackagePart worksheetPart = workbook.getWorksheetPartAt(0);
        reader = new SSMLSheetReader(worksheetPart, workbook);
        testSheetReader();
    }
}
