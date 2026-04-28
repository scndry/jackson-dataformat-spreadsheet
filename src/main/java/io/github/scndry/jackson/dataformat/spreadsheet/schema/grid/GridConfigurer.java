package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;

/**
 * Configurer for sheet-level features: freeze pane, auto filter, and conditional formatting.
 * Conditional formatting rules reference style names declared in
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder}.
 */
public final class GridConfigurer {

    private int _freezePaneColSplit = -1;
    private int _freezePaneRowSplit = -1;
    private boolean _autoFilter;
    private final List<ConditionalFormattingRuleSpec> _rules = new ArrayList<>();

    public GridConfigurer freezePane(final int colSplit, final int rowSplit) {
        _freezePaneColSplit = colSplit;
        _freezePaneRowSplit = rowSplit;
        return this;
    }

    public GridConfigurer autoFilter() {
        _autoFilter = true;
        return this;
    }

    public ConditionalFormattingRuleSpec conditionalFormatting() {
        return new ConditionalFormattingRuleSpec(this);
    }

    void _addRule(final ConditionalFormattingRuleSpec spec) {
        _rules.add(spec);
    }

    public void apply(final Sheet sheet, final Styles styles,
            final SpreadsheetSchema schema, final int lastRow) {
        _applyFreezePane(sheet);
        _applyAutoFilter(sheet, schema, lastRow);
        _applyConditionalFormattings(sheet, styles, schema, lastRow);
    }

    private void _applyFreezePane(final Sheet sheet) {
        if (_freezePaneColSplit >= 0 || _freezePaneRowSplit >= 0) {
            sheet.createFreezePane(
                    Math.max(_freezePaneColSplit, 0),
                    Math.max(_freezePaneRowSplit, 0));
        }
    }

    private void _applyAutoFilter(final Sheet sheet, final SpreadsheetSchema schema, final int lastRow) {
        if (!_autoFilter || schema.columnCount() == 0) return;
        final int firstCol = schema.getOriginColumn();
        final int lastCol = firstCol + schema.columnCount() - 1;
        final int firstRow = schema.getOriginRow();
        final int endRow = lastRow < 0
                ? sheet.getWorkbook().getSpreadsheetVersion().getMaxRows() - 1
                : lastRow;
        sheet.setAutoFilter(new CellRangeAddress(firstRow, endRow, firstCol, lastCol));
    }

    private void _applyConditionalFormattings(final Sheet sheet, final Styles styles,
            final SpreadsheetSchema schema, final int lastRow) {
        if (_rules.isEmpty()) return;
        final Workbook wb = sheet.getWorkbook();
        final int dataRow = schema.getDataRow();
        final int endRow = lastRow < 0 ? wb.getSpreadsheetVersion().getMaxRows() - 1 : lastRow;
        if (dataRow > endRow) return;
        final SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();

        for (final ConditionalFormattingRuleSpec spec : _rules) {
            final int colIndex = schema.columnIndexByName(spec._columnName);
            if (colIndex < 0) {
                throw new IllegalArgumentException("Column '" + spec._columnName
                        + "' not found in schema. Available columns: " + schema.columnNames());
            }
            final CellStyle cs = styles.getStyle(spec._styleName);
            if (cs == null) {
                throw new IllegalArgumentException("Style '" + spec._styleName + "' not found");
            }
            final ConditionalFormattingRule rule = spec._formula2 != null
                    ? scf.createConditionalFormattingRule(spec._operator, spec._formula1, spec._formula2)
                    : scf.createConditionalFormattingRule(spec._operator, spec._formula1);
            _applyCellStyleToDxf(rule, cs, wb);
            scf.addConditionalFormatting(
                    new CellRangeAddress[]{new CellRangeAddress(dataRow, endRow, colIndex, colIndex)}, rule);
        }
    }

    private static void _applyCellStyleToDxf(final ConditionalFormattingRule rule,
            final CellStyle cs, final Workbook wb) {
        if (cs.getFillPattern() != FillPatternType.NO_FILL) {
            final PatternFormatting pf = rule.createPatternFormatting();
            pf.setFillPattern(cs.getFillPattern().getCode());
            if (cs.getFillForegroundColorColor() != null) {
                pf.setFillForegroundColor(cs.getFillForegroundColorColor());
            }
        }
        final Font font = wb.getFontAt(cs.getFontIndex());
        final boolean isXssf = wb.getSpreadsheetVersion() == SpreadsheetVersion.EXCEL2007;
        final boolean hasColor = isXssf
                ? ((XSSFFont) font).getXSSFColor() != null
                : font.getColor() != 0;
        final boolean hasUnderline = font.getUnderline() != Font.U_NONE;
        final boolean hasHeight = font.getFontHeightInPoints() != XSSFFont.DEFAULT_FONT_SIZE;
        if (font.getBold() || font.getItalic() || hasColor || hasUnderline || hasHeight) {
            final FontFormatting ff = rule.createFontFormatting();
            ff.setFontStyle(font.getItalic(), font.getBold());
            if (hasColor) {
                if (isXssf) {
                    ff.setFontColor(((XSSFFont) font).getXSSFColor());
                } else {
                    ff.setFontColorIndex(font.getColor());
                }
            }
            if (hasUnderline) {
                ff.setUnderlineType(font.getUnderline());
            }
            if (hasHeight) {
                ff.setFontHeight(font.getFontHeight());
            }
        }
        if (cs.getBorderLeft() != BorderStyle.NONE || cs.getBorderRight() != BorderStyle.NONE
                || cs.getBorderTop() != BorderStyle.NONE || cs.getBorderBottom() != BorderStyle.NONE) {
            final BorderFormatting bf = rule.createBorderFormatting();
            bf.setBorderLeft(cs.getBorderLeft());
            bf.setBorderRight(cs.getBorderRight());
            bf.setBorderTop(cs.getBorderTop());
            bf.setBorderBottom(cs.getBorderBottom());
            if (isXssf) {
                _applyBorderColor(bf, (XSSFCellStyle) cs);
            } else {
                _applyBorderColor(bf, cs);
            }
        }
    }

    private static void _applyBorderColor(final BorderFormatting bf, final XSSFCellStyle xcs) {
        if (xcs.getLeftBorderXSSFColor() != null) bf.setLeftBorderColor(xcs.getLeftBorderXSSFColor());
        if (xcs.getRightBorderXSSFColor() != null) bf.setRightBorderColor(xcs.getRightBorderXSSFColor());
        if (xcs.getTopBorderXSSFColor() != null) bf.setTopBorderColor(xcs.getTopBorderXSSFColor());
        if (xcs.getBottomBorderXSSFColor() != null) bf.setBottomBorderColor(xcs.getBottomBorderXSSFColor());
    }

    private static void _applyBorderColor(final BorderFormatting bf, final CellStyle cs) {
        if (cs.getLeftBorderColor() != 0) bf.setLeftBorderColor(cs.getLeftBorderColor());
        if (cs.getRightBorderColor() != 0) bf.setRightBorderColor(cs.getRightBorderColor());
        if (cs.getTopBorderColor() != 0) bf.setTopBorderColor(cs.getTopBorderColor());
        if (cs.getBottomBorderColor() != 0) bf.setBottomBorderColor(cs.getBottomBorderColor());
    }
}
