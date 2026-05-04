package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.Incubating;
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
    private final List<ColumnRule> _columnRules = new ArrayList<>();

    private static final class ColumnRule {
        final String column;
        final ConditionalFormatRule rule;
        ColumnRule(final String column, final ConditionalFormatRule rule) {
            this.column = column;
            this.rule = rule;
        }
    }

    public GridConfigurer freezePane(final int colSplit, final int rowSplit) {
        _freezePaneColSplit = colSplit;
        _freezePaneRowSplit = rowSplit;
        return this;
    }

    public GridConfigurer autoFilter() {
        _autoFilter = true;
        return this;
    }

    /**
     * Add one or more conditional formatting rules to a column. Rules apply to the
     * data range of the named column (data start row through last written row).
     *
     * <p>Rules originate from {@link ConditionalFormats} factories, e.g.:
     * <pre>{@code
     * import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.*;
     *
     * grid.conditionalFormatting("score",
     *         greaterThanOrEqual(80).style("good"),
     *         lessThan(60).style("bad"))
     *     .conditionalFormatting("revenue",
     *         colorScale(0, 50_000, 100_000));
     * }</pre>
     *
     * <p>The {@code rule, rules...} shape requires at least one rule at compile time
     * (idiom shared with {@code Path.of(String, String...)} and
     * {@code EnumSet.of(E, E...)}).
     *
     * @param column column name as declared by the data class
     * @param rule  first rule (required)
     * @param rules additional rules for the same column (optional)
     * @return this {@code GridConfigurer} for chaining
     * @since 1.6.0
     */
    @Incubating
    public GridConfigurer conditionalFormatting(
            final String column,
            final ConditionalFormatRule rule,
            final ConditionalFormatRule... rules) {
        Objects.requireNonNull(column, "column");
        Objects.requireNonNull(rule, "rule");
        Objects.requireNonNull(rules, "rules");
        _columnRules.add(new ColumnRule(column, rule));
        for (final ConditionalFormatRule r : rules) {
            Objects.requireNonNull(r, "rules element");
            _columnRules.add(new ColumnRule(column, r));
        }
        return this;
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
        if (_columnRules.isEmpty()) return;
        final Workbook wb = sheet.getWorkbook();
        final int dataRow = schema.getDataRow();
        final int endRow = lastRow < 0 ? wb.getSpreadsheetVersion().getMaxRows() - 1 : lastRow;
        if (dataRow > endRow) return;
        final SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();

        for (final ColumnRule cr : _columnRules) {
            final int colIndex = schema.columnIndexByName(cr.column);
            if (colIndex < 0) {
                throw new IllegalArgumentException("Column '" + cr.column
                        + "' not found in schema. Available columns: " + schema.columnNames());
            }
            final ConditionalFormattingRule rule = _createPoiRule(scf, cr.rule, schema, wb, styles);
            scf.addConditionalFormatting(
                    new CellRangeAddress[]{new CellRangeAddress(dataRow, endRow, colIndex, colIndex)}, rule);
        }
    }

    private static ConditionalFormattingRule _createPoiRule(
            final SheetConditionalFormatting scf,
            final ConditionalFormatRule rule,
            final SpreadsheetSchema schema,
            final Workbook wb,
            final Styles styles) {
        if (rule instanceof CellIsRule) {
            final CellIsRule r = (CellIsRule) rule;
            final String f1 = _resolveOperand(r._operand1, schema);
            final String f2 = r._operand2 != null ? _resolveOperand(r._operand2, schema) : null;
            final ConditionalFormattingRule poiRule = f2 != null
                    ? scf.createConditionalFormattingRule(r._operator, f1, f2)
                    : scf.createConditionalFormattingRule(r._operator, f1);
            _applyStyle(poiRule, r._styleName, styles, wb);
            return poiRule;
        }
        if (rule instanceof ExpressionRule) {
            final ExpressionRule r = (ExpressionRule) rule;
            final ConditionalFormattingRule poiRule = scf.createConditionalFormattingRule(r._formula);
            _applyStyle(poiRule, r._styleName, styles, wb);
            return poiRule;
        }
        if (rule instanceof ColorScaleRule) {
            final ColorScaleRule r = (ColorScaleRule) rule;
            final ConditionalFormattingRule poiRule = scf.createConditionalFormattingColorScaleRule();
            _applyColorScale(poiRule, r);
            return poiRule;
        }
        throw new IllegalStateException("Unknown rule type: " + rule.getClass().getName());
    }

    private static void _applyStyle(final ConditionalFormattingRule rule,
            final String styleName, final Styles styles, final Workbook wb) {
        final CellStyle cs = styles.getStyle(styleName);
        if (cs == null) {
            throw new IllegalArgumentException("Style '" + styleName + "' not found");
        }
        _applyCellStyleToDxf(rule, cs, wb);
    }

    private static void _applyColorScale(final ConditionalFormattingRule poiRule,
            final ColorScaleRule r) {
        // POI's createConditionalFormattingColorScaleRule() pre-populates 3 thresholds
        // (MIN / PERCENTILE 50 / MAX) and 3 empty color slots. Mutate in place (POI's
        // standard pattern in poi-examples/ConditionalFormats).
        final ColorScaleFormatting csf = poiRule.getColorScaleFormatting();
        if (!r._useExcelDefaults) {
            final ConditionalFormattingThreshold[] t = csf.getThresholds();
            t[0].setRangeType(RangeType.NUMBER);
            t[0].setValue(r._min);
            t[1].setRangeType(RangeType.NUMBER);
            t[1].setValue(r._mid);
            t[2].setRangeType(RangeType.NUMBER);
            t[2].setValue(r._max);
        }
        // Excel's standard 3-color scale: red → yellow → green.
        final Color[] colors = csf.getColors();
        ((ExtendedColor) colors[0]).setARGBHex("FFF8696B");
        ((ExtendedColor) colors[1]).setARGBHex("FFFFEB84");
        ((ExtendedColor) colors[2]).setARGBHex("FF63BE7B");
    }

    private static String _resolveOperand(final Object operand, final SpreadsheetSchema schema) {
        if (operand instanceof Formula) {
            final Formula f = (Formula) operand;
            switch (f.kind()) {
                case OF:
                    return f.value();
                case COLUMN:
                    final int col = schema.columnIndexByName(f.value());
                    if (col < 0) {
                        throw new IllegalArgumentException("Formula column '" + f.value()
                                + "' not found in schema. Available columns: " + schema.columnNames());
                    }
                    return "$" + CellReference.convertNumToColString(col) + (schema.getDataRow() + 1);
                default:
                    throw new IllegalStateException("Unknown Formula kind: " + f.kind());
            }
        }
        return (String) operand;
    }

    // Each rule allocates its own dxf — POI does not deduplicate across rules sharing the same style.
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
