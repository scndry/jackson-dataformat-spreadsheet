package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.File;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.junit.jupiter.api.Test;

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
        SSMLWorkbook workbook = SSMLWorkbook.create(strictSource);
        PackagePart part = workbook.getWorksheetPartAt(0);
        reader = new SSMLSheetReader(part, workbook, false, false);
        testSheetReader();
    }
}
