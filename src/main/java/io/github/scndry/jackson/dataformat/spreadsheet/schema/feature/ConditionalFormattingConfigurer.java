package io.github.scndry.jackson.dataformat.spreadsheet.schema.feature;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;

/**
 * Configurer for conditional formatting rules.
 * Rules reference style names declared in
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder}.
 */
public final class ConditionalFormattingConfigurer {

    private final List<Rule> _rules = new ArrayList<>();

    public Rule rule() {
        return new Rule(this);
    }

    void _addRule(final Rule rule) {
        _rules.add(rule);
    }

    public void apply(final Sheet sheet, final Styles styles,
            final SpreadsheetSchema schema, final int lastRow) {
        if (_rules.isEmpty()) return;
        final Workbook wb = sheet.getWorkbook();
        final int dataRow = schema.getDataRow();
        final int endRow = lastRow < 0 ? wb.getSpreadsheetVersion().getMaxRows() - 1 : lastRow;
        final SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();

        for (final Rule cfRule : _rules) {
            final int colIndex = schema.columnIndexByName(cfRule._columnName);
            if (colIndex < 0) {
                throw new IllegalArgumentException("Column '" + cfRule._columnName
                        + "' not found in schema. Available columns: " + schema.columnNames());
            }
            final CellStyle cs = styles.getStyle(cfRule._styleName);
            if (cs == null) {
                throw new IllegalArgumentException("Style '" + cfRule._styleName + "' not found");
            }
            final ConditionalFormattingRule rule = cfRule._formula2 != null
                    ? scf.createConditionalFormattingRule(cfRule._operator, cfRule._formula1, cfRule._formula2)
                    : scf.createConditionalFormattingRule(cfRule._operator, cfRule._formula1);
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
        final boolean hasColor = font instanceof XSSFFont
                && ((XSSFFont) font).getXSSFColor() != null;
        final boolean hasUnderline = font.getUnderline() != Font.U_NONE;
        if (font.getBold() || font.getItalic() || hasColor || hasUnderline) {
            final FontFormatting ff = rule.createFontFormatting();
            ff.setFontStyle(font.getItalic(), font.getBold());
            if (hasColor) {
                ff.setFontColor(((XSSFFont) font).getXSSFColor());
            }
            if (hasUnderline) {
                ff.setUnderlineType(font.getUnderline());
            }
        }
        if (cs.getBorderLeft() != BorderStyle.NONE || cs.getBorderRight() != BorderStyle.NONE
                || cs.getBorderTop() != BorderStyle.NONE || cs.getBorderBottom() != BorderStyle.NONE) {
            final BorderFormatting bf = rule.createBorderFormatting();
            bf.setBorderLeft(cs.getBorderLeft());
            bf.setBorderRight(cs.getBorderRight());
            bf.setBorderTop(cs.getBorderTop());
            bf.setBorderBottom(cs.getBorderBottom());
            if (cs instanceof XSSFCellStyle) {
                final XSSFCellStyle xcs = (XSSFCellStyle) cs;
                _setBorderColorIfPresent(bf, xcs);
            }
        }
    }

    private static void _setBorderColorIfPresent(final BorderFormatting bf,
            final XSSFCellStyle xcs) {
        final XSSFColor left = xcs.getLeftBorderXSSFColor();
        if (left != null) bf.setLeftBorderColor(left);
        final XSSFColor right = xcs.getRightBorderXSSFColor();
        if (right != null) bf.setRightBorderColor(right);
        final XSSFColor top = xcs.getTopBorderXSSFColor();
        if (top != null) bf.setTopBorderColor(top);
        final XSSFColor bottom = xcs.getBottomBorderXSSFColor();
        if (bottom != null) bf.setBottomBorderColor(bottom);
    }

    public static final class Rule {

        private final ConditionalFormattingConfigurer _parent;
        private String _columnName;
        private byte _operator;
        private String _formula1;
        private String _formula2;
        private String _styleName;

        Rule(final ConditionalFormattingConfigurer parent) {
            _parent = parent;
        }

        public Rule column(final String name) {
            _columnName = name;
            return this;
        }

        public Rule greaterThan(final String formula) {
            _operator = ComparisonOperator.GT;
            _formula1 = formula;
            return this;
        }

        public Rule greaterThanOrEqual(final String formula) {
            _operator = ComparisonOperator.GE;
            _formula1 = formula;
            return this;
        }

        public Rule lessThan(final String formula) {
            _operator = ComparisonOperator.LT;
            _formula1 = formula;
            return this;
        }

        public Rule lessThanOrEqual(final String formula) {
            _operator = ComparisonOperator.LE;
            _formula1 = formula;
            return this;
        }

        public Rule equalTo(final String formula) {
            _operator = ComparisonOperator.EQUAL;
            _formula1 = formula;
            return this;
        }

        public Rule notEqualTo(final String formula) {
            _operator = ComparisonOperator.NOT_EQUAL;
            _formula1 = formula;
            return this;
        }

        public Rule between(final String formula1, final String formula2) {
            _operator = ComparisonOperator.BETWEEN;
            _formula1 = formula1;
            _formula2 = formula2;
            return this;
        }

        public Rule notBetween(final String formula1, final String formula2) {
            _operator = ComparisonOperator.NOT_BETWEEN;
            _formula1 = formula1;
            _formula2 = formula2;
            return this;
        }

        public Rule style(final String name) {
            _styleName = name;
            return this;
        }

        public ConditionalFormattingConfigurer end() {
            if (_columnName == null) {
                throw new IllegalStateException("column() must be called before end()");
            }
            if (_formula1 == null) {
                throw new IllegalStateException("A comparison method (greaterThan, equalTo, etc.) must be called before end()");
            }
            if (_styleName == null) {
                throw new IllegalStateException("style() must be called before end()");
            }
            _parent._addRule(this);
            return _parent;
        }
    }
}
