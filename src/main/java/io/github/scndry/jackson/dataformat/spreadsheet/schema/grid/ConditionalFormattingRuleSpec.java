package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.ComparisonOperator;

/**
 * Fluent specification for a single conditional formatting rule.
 * References a column by name and a style by name (defined in
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder}).
 *
 * <p>Operands accept typed values (numeric, boolean, string, date) or
 * {@link Formula} for cell references and arbitrary Excel formula expressions.
 * String operands are auto-quoted for Excel string literal syntax.
 *
 * @see GridConfigurer#conditionalFormatting()
 */
public final class ConditionalFormattingRuleSpec {

    enum Type { CELL_IS, EXPRESSION }

    private final GridConfigurer _parent;
    String _columnName;
    Type _type;
    // CELL_IS
    byte _operator;
    Object _operand1;
    Object _operand2;
    // EXPRESSION
    String _expression;

    String _styleName;

    ConditionalFormattingRuleSpec(final GridConfigurer parent) {
        _parent = parent;
    }

    public ConditionalFormattingRuleSpec column(final String name) {
        _columnName = name;
        return this;
    }

    // ─── greaterThan ───────────────────────────────────────────────

    public ConditionalFormattingRuleSpec greaterThan(final Number value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThan(final LocalDate value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThan(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThan(final Date value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThan(final Calendar value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThan(final Formula formula) {
        return _cellIs(ComparisonOperator.GT, formula);
    }

    // ─── greaterThanOrEqual ────────────────────────────────────────

    public ConditionalFormattingRuleSpec greaterThanOrEqual(final Number value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThanOrEqual(final LocalDate value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThanOrEqual(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThanOrEqual(final Date value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThanOrEqual(final Calendar value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public ConditionalFormattingRuleSpec greaterThanOrEqual(final Formula formula) {
        return _cellIs(ComparisonOperator.GE, formula);
    }

    // ─── lessThan ──────────────────────────────────────────────────

    public ConditionalFormattingRuleSpec lessThan(final Number value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThan(final LocalDate value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThan(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThan(final Date value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThan(final Calendar value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThan(final Formula formula) {
        return _cellIs(ComparisonOperator.LT, formula);
    }

    // ─── lessThanOrEqual ───────────────────────────────────────────

    public ConditionalFormattingRuleSpec lessThanOrEqual(final Number value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThanOrEqual(final LocalDate value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThanOrEqual(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThanOrEqual(final Date value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThanOrEqual(final Calendar value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public ConditionalFormattingRuleSpec lessThanOrEqual(final Formula formula) {
        return _cellIs(ComparisonOperator.LE, formula);
    }

    // ─── equalTo ───────────────────────────────────────────────────

    public ConditionalFormattingRuleSpec equalTo(final Number value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec equalTo(final boolean value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec equalTo(final String value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec equalTo(final LocalDate value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec equalTo(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec equalTo(final Date value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec equalTo(final Calendar value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec equalTo(final Formula formula) {
        return _cellIs(ComparisonOperator.EQUAL, formula);
    }

    // ─── notEqualTo ────────────────────────────────────────────────

    public ConditionalFormattingRuleSpec notEqualTo(final Number value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec notEqualTo(final boolean value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec notEqualTo(final String value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec notEqualTo(final LocalDate value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec notEqualTo(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec notEqualTo(final Date value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec notEqualTo(final Calendar value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public ConditionalFormattingRuleSpec notEqualTo(final Formula formula) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, formula);
    }

    // ─── between ───────────────────────────────────────────────────

    public ConditionalFormattingRuleSpec between(final Number low, final Number high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec between(final LocalDate low, final LocalDate high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec between(final LocalDateTime low, final LocalDateTime high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec between(final Date low, final Date high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec between(final Calendar low, final Calendar high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec between(final Formula low, final Formula high) {
        return _cellIs(ComparisonOperator.BETWEEN, low, high);
    }

    // ─── notBetween ────────────────────────────────────────────────

    public ConditionalFormattingRuleSpec notBetween(final Number low, final Number high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec notBetween(final LocalDate low, final LocalDate high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec notBetween(final LocalDateTime low, final LocalDateTime high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec notBetween(final Date low, final Date high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec notBetween(final Calendar low, final Calendar high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public ConditionalFormattingRuleSpec notBetween(final Formula low, final Formula high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, low, high);
    }

    // ─── expression rule (type="expression") ───────────────────────

    /**
     * Arbitrary boolean Excel formula evaluated per cell in the conditional formatting range.
     * Maps to OOXML {@code <cfRule type="expression">}. The formula must return TRUE or FALSE.
     *
     * <pre>{@code
     * .column("price").expression("=AND($A1>0, $B1<100)").style("warn")
     * }</pre>
     */
    public ConditionalFormattingRuleSpec expression(final String formula) {
        if (formula == null || formula.isEmpty()) {
            throw new IllegalArgumentException("Expression formula must not be empty");
        }
        _type = Type.EXPRESSION;
        _expression = formula;
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
        if (_type == null) {
            throw new IllegalStateException(
                    "A comparison method (greaterThan, equalTo, etc.) or expression() must be called before end()");
        }
        if (_styleName == null) {
            throw new IllegalStateException("style() must be called before end()");
        }
        _parent._addRule(this);
        return _parent;
    }

    private ConditionalFormattingRuleSpec _cellIs(final byte operator, final Object operand) {
        _type = Type.CELL_IS;
        _operator = operator;
        _operand1 = operand;
        _operand2 = null;
        return this;
    }

    private ConditionalFormattingRuleSpec _cellIs(final byte operator, final Object low, final Object high) {
        _type = Type.CELL_IS;
        _operator = operator;
        _operand1 = low;
        _operand2 = high;
        return this;
    }

    // ─── formatters ────────────────────────────────────────────────

    private static String _f(final Number value) {
        return value.toString();
    }

    private static String _f(final boolean value) {
        return value ? "TRUE" : "FALSE";
    }

    private static String _f(final String value) {
        // Excel string literal: surrounded by quotes; internal quotes doubled
        return "\"" + value.replace("\"", "\"\"") + "\"";
    }

    private static String _f(final LocalDate value) {
        return "DATE(" + value.getYear() + ","
                + value.getMonthValue() + ","
                + value.getDayOfMonth() + ")";
    }

    private static String _f(final LocalDateTime value) {
        return _f(value.toLocalDate())
                + "+TIME(" + value.getHour() + ","
                + value.getMinute() + ","
                + value.getSecond() + ")";
    }

    private static String _f(final Date value) {
        return _f(LocalDateTime.ofInstant(value.toInstant(), ZoneId.systemDefault()));
    }

    private static String _f(final Calendar value) {
        return _f(LocalDateTime.ofInstant(value.toInstant(), value.getTimeZone().toZoneId()));
    }
}
