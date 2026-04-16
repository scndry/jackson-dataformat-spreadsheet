package io.github.scndry.jackson.dataformat.spreadsheet.poi.ss;

import java.io.File;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReaderTestBase;

class POISheetReaderTest extends SheetReaderTestBase {

    @Test
    void readViaPoiObjectModel() throws Exception {
        File file = createHeadlessFixture();
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        reader = new POISheetReader(sheet);
        testSheetReader();
    }
}
