package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.ColorConfigurer.FontColor;
import org.apache.poi.common.usermodel.fonts.FontCharset;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.function.Function;

public final class FontBuilder extends NestedBuilder<CellStyleBuilder> implements Builder<Font> {

    private final FontColor _color = new FontColor(XSSFFont::setColor, HSSFFont::setColor);
    private String _fontName = XSSFFont.DEFAULT_FONT_NAME;
    private short _fontHeight = XSSFFont.DEFAULT_FONT_SIZE;
    private boolean _italic;
    private boolean _strikeout;
    private short _typeOffset;
    private FontUnderline _underline = FontUnderline.NONE;
    private FontCharset _charSet = FontCharset.ANSI;
    private boolean _bold;

    @Override
    public FontBuilder copy() {
        return new FontBuilder()
                .fontName(_fontName)
                .fontHeight(_fontHeight)
                .italic(_italic)
                .strikeout(_strikeout)
                .typeOffset(_typeOffset)
                .underline(_underline)
                .charSet(_charSet)
                .bold(_bold)
                .color(_color.rgb);
    }

    @Override
    public Font build(Workbook workbook) {
        final Font font = workbook.createFont();
        font.setFontName(_fontName);
        font.setFontHeightInPoints(_fontHeight);
        font.setItalic(_italic);
        font.setStrikeout(_strikeout);
        font.setTypeOffset(_typeOffset);
        font.setBold(_bold);
        font.setUnderline(_underline.getByteValue());
        font.setCharSet(_charSet.getNativeId());
        _color.build(workbook).accept(font);
        return font;
    }

    //@formatter:off
    public FontBuilder fontName(String fontName) { _fontName = fontName; return this; }
    public FontBuilder fontHeight(short fontHeight) { _fontHeight = fontHeight; return this; }
    public FontBuilder italic() { return italic(true); }
    public FontBuilder italic(boolean italic) { _italic = italic; return this; }
    public FontBuilder strikeout() { return strikeout(true); }
    public FontBuilder strikeout(boolean strikeout) { _strikeout = strikeout; return this; }
    public FontBuilder typeOffset(short typeOffset) { _typeOffset = typeOffset; return this; }
    public FontBuilder bold() { return bold(true); }
    public FontBuilder bold(boolean bold) { _bold = bold; return this; }
    FontBuilder color(byte[] color) { _color.rgb = color; return this; }
    FontBuilder underline(FontUnderline underline) { _underline = underline; return this; }
    FontBuilder charSet(FontCharset charSet) { _charSet = charSet; return this; }
    public FontBuilder color(int rgb) { return color((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public FontBuilder color(int r, int g, int b) { return color(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public FontBuilder color(IndexedColors color) { return color(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public FontUnderlineConfigurer underline() { return new FontUnderlineConfigurer(this::underline); }
    public FontCharsetConfigurer charSet() { return new FontCharsetConfigurer(this::charSet); }

    public static class FontUnderlineConfigurer extends PropertyConfigurer<FontUnderline, FontBuilder> {
        FontUnderlineConfigurer(Function<FontUnderline, FontBuilder> property) { super(property); }
        public FontBuilder single() { return apply(FontUnderline.SINGLE); }
        public FontBuilder doubled() { return apply(FontUnderline.DOUBLE); }
        public FontBuilder singleAccounting() { return apply(FontUnderline.SINGLE_ACCOUNTING); }
        public FontBuilder doubleAccounting() { return apply(FontUnderline.DOUBLE_ACCOUNTING); }
        public FontBuilder none() { return apply(FontUnderline.NONE); }
    }
    public static class FontCharsetConfigurer extends PropertyConfigurer<FontCharset, FontBuilder> {
        FontCharsetConfigurer(Function<FontCharset, FontBuilder> property) { super(property); }
        public FontBuilder ansi() { return apply(FontCharset.ANSI); }
        public FontBuilder defaults() { return apply(FontCharset.DEFAULT); }
        public FontBuilder symbol() { return apply(FontCharset.SYMBOL); }
        public FontBuilder mac() { return apply(FontCharset.MAC); }
        public FontBuilder shiftjis() { return apply(FontCharset.SHIFTJIS); }
        public FontBuilder hangul() { return apply(FontCharset.HANGUL); }
        public FontBuilder johab() { return apply(FontCharset.JOHAB); }
        public FontBuilder gb2312() { return apply(FontCharset.GB2312); }
        public FontBuilder chinesebig5() { return apply(FontCharset.CHINESEBIG5); }
        public FontBuilder greek() { return apply(FontCharset.GREEK); }
        public FontBuilder turkish() { return apply(FontCharset.TURKISH); }
        public FontBuilder vietnamese() { return apply(FontCharset.VIETNAMESE); }
        public FontBuilder hebrew() { return apply(FontCharset.HEBREW); }
        public FontBuilder arabic() { return apply(FontCharset.ARABIC); }
        public FontBuilder baltic() { return apply(FontCharset.BALTIC); }
        public FontBuilder russian() { return apply(FontCharset.RUSSIAN); }
        public FontBuilder thai() { return apply(FontCharset.THAI); }
        public FontBuilder easteurope() { return apply(FontCharset.EASTEUROPE); }
        public FontBuilder oem() { return apply(FontCharset.OEM); }
    }
    //@formatter:on
}
