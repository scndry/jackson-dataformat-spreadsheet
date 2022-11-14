package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

interface ColorBinder<T, C extends Color> {

    void bind(T t, C c);

    C color(Workbook workbook, byte[] rgb);

    interface XSSFColorBinder<T> extends ColorBinder<T, XSSFColor> {
        @Override
        void bind(T t, XSSFColor c);

        default XSSFColor color(final Workbook workbook, final byte[] rgb) {
            final XSSFColor color = (XSSFColor) workbook.getCreationHelper().createExtendedColor();
            color.setRGB(rgb);
            return color;
        }
    }

    interface HSSFColorBinder<T> extends ColorBinder<T, HSSFColor> {
        @Override
        default void bind(final T t, final HSSFColor c) {
            if (c != null) {
                accept(t, c.getIndex());
            }
        }

        void accept(T t, short c);

        default HSSFColor color(final Workbook workbook, final byte[] rgb) {
            final HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            final HSSFColor color = palette.findColor(rgb[0], rgb[1], rgb[2]);
            if (color == null) {
                return palette.findSimilarColor(rgb[0], rgb[1], rgb[2]);
            }
            return color;
        }
    }
}
