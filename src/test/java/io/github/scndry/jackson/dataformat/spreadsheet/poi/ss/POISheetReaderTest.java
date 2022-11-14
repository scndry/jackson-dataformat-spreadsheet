package io.github.scndry.jackson.dataformat.spreadsheet.poi.ss;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReaderTestBase;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class POISheetReaderTest extends SheetReaderTestBase {

    @Test
    void test() throws Exception {
        final Workbook workbook = new XSSFWorkbook(source);
        final Sheet sheet = workbook.getSheetAt(0);
        reader = new POISheetReader(sheet);
        testSheetReader();
    }
}
