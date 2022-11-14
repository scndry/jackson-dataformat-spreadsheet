package io.github.scndry.jackson.dataformat.spreadsheet.schema.style;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.ColorConfigurer.CellStyleColor;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.function.Function;

public final class CellStyleBuilder extends NestedBuilder<StylesBuilder> implements Builder<CellStyle> {

    public static final String CONVENIENCE_FORMAT_GENERAL = "General";
    public static final String CONVENIENCE_FORMAT_TEXT = "@";
    public static final String CONVENIENCE_FORMAT_NUMBER_INT = "0";
    public static final String CONVENIENCE_FORMAT_NUMBER_FLOAT = "0.00";
    public static final String CONVENIENCE_FORMAT_NUMBER_INT_WITH_COMMA = "#,##0";
    public static final String CONVENIENCE_FORMAT_NUMBER_FLOAT_WITH_COMMA = "#,##0.00";
    public static final String CONVENIENCE_FORMAT_DATE = "yyyy-mm-dd";
    public static final String CONVENIENCE_FORMAT_DATE_TIME = "yyyy-mm-dd hh:mm:ss";

    private final CellStyleColor _leftBorderColor = new CellStyleColor(XSSFCellStyle::setLeftBorderColor, HSSFCellStyle::setLeftBorderColor);
    private final CellStyleColor _rightBorderColor = new CellStyleColor(XSSFCellStyle::setRightBorderColor, HSSFCellStyle::setRightBorderColor);
    private final CellStyleColor _topBorderColor = new CellStyleColor(XSSFCellStyle::setTopBorderColor, HSSFCellStyle::setTopBorderColor);
    private final CellStyleColor _bottomBorderColor = new CellStyleColor(XSSFCellStyle::setBottomBorderColor, HSSFCellStyle::setBottomBorderColor);
    private final CellStyleColor _fillBackgroundColor = new CellStyleColor(XSSFCellStyle::setFillBackgroundColor, HSSFCellStyle::setFillBackgroundColor);
    private final CellStyleColor _fillForegroundColor = new CellStyleColor(XSSFCellStyle::setFillForegroundColor, HSSFCellStyle::setFillForegroundColor);
    private String _dataFormat;
    private boolean _hidden;
    private boolean _locked;
    private boolean _quotePrefixed;
    private HorizontalAlignment _alignment = HorizontalAlignment.GENERAL;
    private boolean _wrapText;
    private VerticalAlignment _verticalAlignment = VerticalAlignment.CENTER;
    private short _rotation;
    private short _indention;
    private BorderStyle _borderLeft = BorderStyle.NONE;
    private BorderStyle _borderRight = BorderStyle.NONE;
    private BorderStyle _borderTop = BorderStyle.NONE;
    private BorderStyle _borderBottom = BorderStyle.NONE;
    private FillPatternType _fillPattern = FillPatternType.NO_FILL;
    private boolean _shrinkToFit;
    private FontBuilder _fontBuilder;

    @Override
    public CellStyleBuilder copy() {
        return new CellStyleBuilder()
                .dataFormat(_dataFormat)
                .hidden(_hidden)
                .locked(_locked)
                .quotePrefixed(_quotePrefixed)
                .alignment(_alignment)
                .wrapText(_wrapText)
                .verticalAlignment(_verticalAlignment)
                .rotation(_rotation)
                .indention(_indention)
                .borderLeft(_borderLeft)
                .borderRight(_borderRight)
                .borderTop(_borderTop)
                .borderBottom(_borderBottom)
                .leftBorderColor(_leftBorderColor.rgb)
                .rightBorderColor(_rightBorderColor.rgb)
                .topBorderColor(_topBorderColor.rgb)
                .bottomBorderColor(_bottomBorderColor.rgb)
                .fillPattern(_fillPattern)
                .fillBackgroundColor(_fillBackgroundColor.rgb)
                .fillForegroundColor(_fillForegroundColor.rgb)
                .shrinkToFit(_shrinkToFit)
                .fontBuilder(_fontBuilder == null ? null : _fontBuilder.copy());
    }

    @Override
    public CellStyle build(Workbook workbook) {
        final CellStyle style = workbook.createCellStyle();
        final CreationHelper helper = workbook.getCreationHelper();
        style.setHidden(_hidden);
        style.setLocked(_locked);
        style.setQuotePrefixed(_quotePrefixed);
        style.setAlignment(_alignment);
        style.setWrapText(_wrapText);
        style.setVerticalAlignment(_verticalAlignment);
        style.setRotation(_rotation);
        style.setIndention(_indention);
        style.setBorderLeft(_borderLeft);
        style.setBorderRight(_borderRight);
        style.setBorderTop(_borderTop);
        style.setBorderBottom(_borderBottom);
        style.setFillPattern(_fillPattern);
        style.setShrinkToFit(_shrinkToFit);
        if (_fontBuilder != null) {
            style.setFont(_fontBuilder.build(workbook));
        }
        if (_dataFormat != null) {
            style.setDataFormat(helper.createDataFormat().getFormat(_dataFormat));
        }
        _leftBorderColor.build(workbook).accept(style);
        _rightBorderColor.build(workbook).accept(style);
        _topBorderColor.build(workbook).accept(style);
        _bottomBorderColor.build(workbook).accept(style);
        _fillBackgroundColor.build(workbook).accept(style);
        _fillForegroundColor.build(workbook).accept(style);
        return style;
    }

    public FontBuilder font() {
        _fontBuilder = new FontBuilder();
        _fontBuilder.setBuilder(this);
        return _fontBuilder;
    }

    //@formatter:off
    public CellStyleBuilder dataFormat(String dataFormat) { _dataFormat = dataFormat; return this; }
    public DataFormatConfigurer dataFormat() { return new DataFormatConfigurer(this::dataFormat); }
    public CellStyleBuilder hidden() { return hidden(true); }
    public CellStyleBuilder hidden(boolean hidden) { _hidden = hidden; return this; }
    public CellStyleBuilder locked() { return locked(true); }
    public CellStyleBuilder locked(boolean locked) { _locked = locked; return this; }
    public CellStyleBuilder quotePrefixed() { return quotePrefixed(true); }
    public CellStyleBuilder quotePrefixed(boolean quotePrefixed) { _quotePrefixed = quotePrefixed; return this; }
    public CellStyleBuilder wrapText() { return wrapText(true); }
    public CellStyleBuilder wrapText(boolean wrapText) { _wrapText = wrapText; return this; }
    public CellStyleBuilder rotation(short rotation) { _rotation = rotation; return this; }
    public CellStyleBuilder indention(short indention) { _indention = indention; return this; }
    public CellStyleBuilder shrinkToFit() { return shrinkToFit(true); }
    public CellStyleBuilder shrinkToFit(boolean shrinkToFit) { _shrinkToFit = shrinkToFit; return this; }
    CellStyleBuilder fontBuilder(FontBuilder fontBuilder) { _fontBuilder = fontBuilder; return this; }
    CellStyleBuilder borderColor(byte[] borderColor) { return leftBorderColor(borderColor).rightBorderColor(borderColor).topBorderColor(borderColor).bottomBorderColor(borderColor); }
    CellStyleBuilder leftBorderColor(byte[] leftBorderColor) { _leftBorderColor.rgb = leftBorderColor; return this; }
    CellStyleBuilder rightBorderColor(byte[] rightBorderColor) { _rightBorderColor.rgb = rightBorderColor; return this; }
    CellStyleBuilder topBorderColor(byte[] topBorderColor) { _topBorderColor.rgb = topBorderColor; return this; }
    CellStyleBuilder bottomBorderColor(byte[] bottomBorderColor) { _bottomBorderColor.rgb = bottomBorderColor; return this; }
    CellStyleBuilder fillBackgroundColor(byte[] fillBackgroundColor) { _fillBackgroundColor.rgb = fillBackgroundColor; return this; }
    CellStyleBuilder fillForegroundColor(byte[] fillForegroundColor) { _fillForegroundColor.rgb = fillForegroundColor; return this; }
    CellStyleBuilder alignment(HorizontalAlignment alignment) { _alignment = alignment; return this; }
    CellStyleBuilder verticalAlignment(VerticalAlignment verticalAlignment) { _verticalAlignment = verticalAlignment; return this; }
    CellStyleBuilder border(BorderStyle border) { return borderLeft(border).borderRight(border).borderTop(border). borderBottom(border); }
    CellStyleBuilder borderLeft(BorderStyle borderLeft) { _borderLeft = borderLeft; return this; }
    CellStyleBuilder borderRight(BorderStyle borderRight) { _borderRight = borderRight; return this; }
    CellStyleBuilder borderTop(BorderStyle borderTop) { _borderTop = borderTop; return this; }
    CellStyleBuilder borderBottom(BorderStyle borderBottom) { _borderBottom = borderBottom; return this; }
    CellStyleBuilder fillPattern(FillPatternType fillPattern) { _fillPattern = fillPattern; return this; }
    public CellStyleBuilder borderColor(int rgb) { return borderColor((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public CellStyleBuilder borderColor(int r, int g, int b) { return borderColor(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public CellStyleBuilder borderColor(IndexedColors color) { return borderColor(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public CellStyleBuilder leftBorderColor(int rgb) { return leftBorderColor((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public CellStyleBuilder leftBorderColor(int r, int g, int b) { return leftBorderColor(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public CellStyleBuilder leftBorderColor(IndexedColors color) { return leftBorderColor(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public CellStyleBuilder rightBorderColor(int rgb) { return rightBorderColor((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public CellStyleBuilder rightBorderColor(int r, int g, int b) { return rightBorderColor(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public CellStyleBuilder rightBorderColor(IndexedColors color) { return rightBorderColor(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public CellStyleBuilder topBorderColor(int rgb) { return topBorderColor((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public CellStyleBuilder topBorderColor(int r, int g, int b) { return topBorderColor(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public CellStyleBuilder topBorderColor(IndexedColors color) { return topBorderColor(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public CellStyleBuilder bottomBorderColor(int rgb) { return bottomBorderColor((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public CellStyleBuilder bottomBorderColor(int r, int g, int b) { return bottomBorderColor(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public CellStyleBuilder bottomBorderColor(IndexedColors color) { return bottomBorderColor(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public CellStyleBuilder fillBackgroundColor(int rgb) { return fillBackgroundColor((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public CellStyleBuilder fillBackgroundColor(int r, int g, int b) { return fillBackgroundColor(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public CellStyleBuilder fillBackgroundColor(IndexedColors color) { return fillBackgroundColor(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public CellStyleBuilder fillForegroundColor(int rgb) { return fillForegroundColor((rgb & 0xFF0000) >> 16, (rgb & 0xFF00) >> 8, rgb & 0xFF); }
    public CellStyleBuilder fillForegroundColor(int r, int g, int b) { return fillForegroundColor(new byte[]{(byte) r, (byte) g, (byte) b}); }
    public CellStyleBuilder fillForegroundColor(IndexedColors color) { return fillForegroundColor(DefaultIndexedColorMap.getDefaultRGB(color.getIndex())); }
    public HorizontalAlignmentConfigurer alignment() { return new HorizontalAlignmentConfigurer(this::alignment); }
    public VerticalAlignmentConfigurer verticalAlignment() { return new VerticalAlignmentConfigurer(this::verticalAlignment); }
    public BorderStyleConfigurer border() { return new BorderStyleConfigurer(this::border); }
    public BorderStyleConfigurer borderLeft() { return new BorderStyleConfigurer(this::borderLeft); }
    public BorderStyleConfigurer borderRight() { return new BorderStyleConfigurer(this::borderRight); }
    public BorderStyleConfigurer borderTop() { return new BorderStyleConfigurer(this::borderTop); }
    public BorderStyleConfigurer borderBottom() { return new BorderStyleConfigurer(this::borderBottom); }
    public FillPatternConfigurer fillPattern() { return new FillPatternConfigurer(this::fillPattern); }
    public static class DataFormatConfigurer extends PropertyConfigurer<String, CellStyleBuilder> {

    DataFormatConfigurer(Function<String, CellStyleBuilder> property) { super(property); }
        public CellStyleBuilder general() { return apply(CONVENIENCE_FORMAT_GENERAL); }
        public CellStyleBuilder text() { return apply(CONVENIENCE_FORMAT_TEXT); }
        public CellStyleBuilder numberInt() { return apply(CONVENIENCE_FORMAT_NUMBER_INT); }
        public CellStyleBuilder numberFloat() { return apply(CONVENIENCE_FORMAT_NUMBER_FLOAT); }
        public CellStyleBuilder numberIntWithComma() { return apply(CONVENIENCE_FORMAT_NUMBER_INT_WITH_COMMA); }
        public CellStyleBuilder numberFloatWithComma() { return apply(CONVENIENCE_FORMAT_NUMBER_FLOAT_WITH_COMMA); }
        public CellStyleBuilder date() { return apply(CONVENIENCE_FORMAT_DATE); }
        public CellStyleBuilder dateTime() { return apply(CONVENIENCE_FORMAT_DATE_TIME); }
    }
    public static class HorizontalAlignmentConfigurer extends PropertyConfigurer<HorizontalAlignment, CellStyleBuilder> {
        HorizontalAlignmentConfigurer(Function<HorizontalAlignment, CellStyleBuilder> property) { super(property); }
        public CellStyleBuilder general() { return apply(HorizontalAlignment.GENERAL); }
        public CellStyleBuilder left() { return apply(HorizontalAlignment.LEFT); }
        public CellStyleBuilder center() { return apply(HorizontalAlignment.CENTER); }
        public CellStyleBuilder right() { return apply(HorizontalAlignment.RIGHT); }
        public CellStyleBuilder fill() { return apply(HorizontalAlignment.FILL); }
        public CellStyleBuilder justify() { return apply(HorizontalAlignment.JUSTIFY); }
        public CellStyleBuilder centerSelection() { return apply(HorizontalAlignment.CENTER_SELECTION); }
        public CellStyleBuilder distributed() { return apply(HorizontalAlignment.DISTRIBUTED); }
    }
    public static class VerticalAlignmentConfigurer extends PropertyConfigurer<VerticalAlignment, CellStyleBuilder> {
        VerticalAlignmentConfigurer(Function<VerticalAlignment, CellStyleBuilder> property) { super(property); }
        public CellStyleBuilder top() { return apply(VerticalAlignment.TOP); }
        public CellStyleBuilder center() { return apply(VerticalAlignment.CENTER); }
        public CellStyleBuilder bottom() { return apply(VerticalAlignment.BOTTOM); }
        public CellStyleBuilder justify() { return apply(VerticalAlignment.JUSTIFY); }
        public CellStyleBuilder distributed() { return apply(VerticalAlignment.DISTRIBUTED); }
    }
    public static class BorderStyleConfigurer extends PropertyConfigurer<BorderStyle, CellStyleBuilder> {
        BorderStyleConfigurer(Function<BorderStyle, CellStyleBuilder> property) { super(property); }
        public CellStyleBuilder none() { return apply(BorderStyle.NONE); }
        public CellStyleBuilder thin() { return apply(BorderStyle.THIN); }
        public CellStyleBuilder medium() { return apply(BorderStyle.MEDIUM); }
        public CellStyleBuilder dashed() { return apply(BorderStyle.DASHED); }
        public CellStyleBuilder dotted() { return apply(BorderStyle.DOTTED); }
        public CellStyleBuilder thick() { return apply(BorderStyle.THICK); }
        public CellStyleBuilder doubled() { return apply(BorderStyle.DOUBLE); }
        public CellStyleBuilder hair() { return apply(BorderStyle.HAIR); }
        public CellStyleBuilder mediumDashed() { return apply(BorderStyle.MEDIUM_DASHED); }
        public CellStyleBuilder dashDot() { return apply(BorderStyle.DASH_DOT); }
        public CellStyleBuilder mediumDashDot() { return apply(BorderStyle.MEDIUM_DASH_DOT); }
        public CellStyleBuilder dashDotDot() { return apply(BorderStyle.DASH_DOT_DOT); }
        public CellStyleBuilder mediumDashDotDot() { return apply(BorderStyle.MEDIUM_DASH_DOT_DOT); }
        public CellStyleBuilder slantedDashDot() { return apply(BorderStyle.SLANTED_DASH_DOT); }
    }
    public static class FillPatternConfigurer extends PropertyConfigurer<FillPatternType, CellStyleBuilder> {
        FillPatternConfigurer(Function<FillPatternType, CellStyleBuilder> property) { super(property); }
        public CellStyleBuilder noFill() { return apply(FillPatternType.NO_FILL); }
        public CellStyleBuilder solidForeground() { return apply(FillPatternType.SOLID_FOREGROUND); }
        public CellStyleBuilder fineDots() { return apply(FillPatternType.FINE_DOTS); }
        public CellStyleBuilder altBars() { return apply(FillPatternType.ALT_BARS); }
        public CellStyleBuilder sparseDots() { return apply(FillPatternType.SPARSE_DOTS); }
        public CellStyleBuilder thickHorzBands() { return apply(FillPatternType.THICK_HORZ_BANDS); }
        public CellStyleBuilder thickVertBands() { return apply(FillPatternType.THICK_VERT_BANDS); }
        public CellStyleBuilder thickBackwardDiag() { return apply(FillPatternType.THICK_BACKWARD_DIAG); }
        public CellStyleBuilder thickForwardDiag() { return apply(FillPatternType.THICK_FORWARD_DIAG); }
        public CellStyleBuilder bigSpots() { return apply(FillPatternType.BIG_SPOTS); }
        public CellStyleBuilder bricks() { return apply(FillPatternType.BRICKS); }
        public CellStyleBuilder thinHorzBands() { return apply(FillPatternType.THIN_HORZ_BANDS); }
        public CellStyleBuilder thinVertBands() { return apply(FillPatternType.THIN_VERT_BANDS); }
        public CellStyleBuilder thinBackwardDiag() { return apply(FillPatternType.THIN_BACKWARD_DIAG); }
        public CellStyleBuilder thinForwardDiag() { return apply(FillPatternType.THIN_FORWARD_DIAG); }
        public CellStyleBuilder squares() { return apply(FillPatternType.SQUARES); }
        public CellStyleBuilder diamonds() { return apply(FillPatternType.DIAMONDS); }
        public CellStyleBuilder lessDots() { return apply(FillPatternType.LESS_DOTS); }
        public CellStyleBuilder leastDots() { return apply(FillPatternType.LEAST_DOTS); }
    }
    //@formatter:on
}
