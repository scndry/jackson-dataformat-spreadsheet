package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.ColorBinder.HSSFColorBinder;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.ColorBinder.XSSFColorBinder;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.function.Consumer;

abstract class ColorConfigurer<T, X extends T, H extends T> {

    protected final XSSFColorBinder<X> _xssf;
    protected final HSSFColorBinder<H> _hssf;
    protected byte[] rgb;

    ColorConfigurer(final XSSFColorBinder<X> xssf, final HSSFColorBinder<H> hssf) {
        _xssf = xssf;
        _hssf = hssf;
    }

    @SuppressWarnings("unchecked")
    public Consumer<T> build(final Workbook workbook) {
        if (rgb == null) {
            return t -> { /* no op */ };
        }
        if (workbook.getSpreadsheetVersion() == SpreadsheetVersion.EXCEL2007) {
            return t -> _xssf.bind((X) t, _xssf.color(workbook, rgb));
        }
        return t -> _hssf.bind((H) t, _hssf.color(workbook, rgb));
    }

    static final class CellStyleColor extends ColorConfigurer<CellStyle, XSSFCellStyle, HSSFCellStyle> {
        public CellStyleColor(final XSSFColorBinder<XSSFCellStyle> xssf, final HSSFColorBinder<HSSFCellStyle> hssf) {
            super(xssf, hssf);
        }
    }

    static final class FontColor extends ColorConfigurer<Font, XSSFFont, HSSFFont> {
        public FontColor(final XSSFColorBinder<XSSFFont> xssf, final HSSFColorBinder<HSSFFont> hssf) {
            super(xssf, hssf);
        }
    }
}
