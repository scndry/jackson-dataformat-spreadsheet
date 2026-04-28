package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import org.apache.poi.ss.usermodel.ComparisonOperator;

/**
 * Fluent specification for a single conditional formatting rule.
 * References a column by name and a style by name (defined in
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder}).
 *
 * @see GridConfigurer#conditionalFormatting()
 */
public final class ConditionalFormattingRuleSpec {

    private final GridConfigurer _parent;
    String _columnName;
    byte _operator;
    String _formula1;
    String _formula2;
    String _styleName;

    ConditionalFormattingRuleSpec(final GridConfigurer parent) {
        _parent = parent;
    }

    public ConditionalFormattingRuleSpec column(final String name) {
        _columnName = name;
        return this;
    }

    public ConditionalFormattingRuleSpec greaterThan(final String formula) {
        _operator = ComparisonOperator.GT;
        _formula1 = formula;
        return this;
    }

    public ConditionalFormattingRuleSpec greaterThanOrEqual(final String formula) {
        _operator = ComparisonOperator.GE;
        _formula1 = formula;
        return this;
    }

    public ConditionalFormattingRuleSpec lessThan(final String formula) {
        _operator = ComparisonOperator.LT;
        _formula1 = formula;
        return this;
    }

    public ConditionalFormattingRuleSpec lessThanOrEqual(final String formula) {
        _operator = ComparisonOperator.LE;
        _formula1 = formula;
        return this;
    }

    public ConditionalFormattingRuleSpec equalTo(final String formula) {
        _operator = ComparisonOperator.EQUAL;
        _formula1 = formula;
        return this;
    }

    public ConditionalFormattingRuleSpec notEqualTo(final String formula) {
        _operator = ComparisonOperator.NOT_EQUAL;
        _formula1 = formula;
        return this;
    }

    public ConditionalFormattingRuleSpec between(final String formula1, final String formula2) {
        _operator = ComparisonOperator.BETWEEN;
        _formula1 = formula1;
        _formula2 = formula2;
        return this;
    }

    public ConditionalFormattingRuleSpec notBetween(final String formula1, final String formula2) {
        _operator = ComparisonOperator.NOT_BETWEEN;
        _formula1 = formula1;
        _formula2 = formula2;
        return this;
    }

    public ConditionalFormattingRuleSpec style(final String name) {
        _styleName = name;
        return this;
    }

    public GridConfigurer end() {
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
