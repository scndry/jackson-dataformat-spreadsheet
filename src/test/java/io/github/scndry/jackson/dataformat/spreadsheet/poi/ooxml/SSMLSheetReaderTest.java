package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReaderTestBase;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.junit.jupiter.api.Test;

class SSMLSheetReaderTest extends SheetReaderTestBase {

    @Test
    void test() throws Exception {
        final SSMLWorkbook workbook = SSMLWorkbook.create(source);
        final PackagePart worksheetPart = workbook.getSheetAt(0);
        reader = new SSMLSheetReader(worksheetPart, workbook);
        testSheetReader();
    }
}
