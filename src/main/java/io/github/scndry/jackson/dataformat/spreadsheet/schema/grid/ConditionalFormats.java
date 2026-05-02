package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.Objects;

import org.apache.poi.ss.usermodel.ComparisonOperator;

/**
 * Catalog of factory methods for conditional formatting rules. Static-import for
 * fluent usage:
 *
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
 * <p>Categories:
 * <ul>
 *   <li><b>cellIs</b> — comparison operators ({@link #greaterThan}, {@link #lessThan},
 *       {@link #between}, {@link #equalTo}, etc.) with typed value overloads. Returns
 *       {@link FormatCondition}; complete with {@link FormatCondition#style(String)}.</li>
 *   <li><b>expression</b> — arbitrary boolean Excel formula via {@link #expression}.
 *       Returns {@link FormatCondition}.</li>
 *   <li><b>visualization</b> — {@link #colorScale()} / {@link #colorScale(double, double, double)}
 *       returns {@link ConditionalFormatRule} directly (no styling required).</li>
 *   <li><b>operand factories</b> — {@link #formula(String)} (raw passthrough) and
 *       {@link #columnRef(String)} (schema-aware row-relative reference) produce
 *       {@link Formula} operands for cellIs comparisons.</li>
 * </ul>
 *
 * <p>Methods that accept a typed value ({@link Number} / {@link LocalDate} /
 * {@link LocalDateTime} / {@link Date} / {@link Calendar} / {@link String} / {@code boolean})
 * auto-convert to the appropriate Excel formula representation. Use {@link #formula(String)}
 * or {@link #columnRef(String)} for raw formulas or column-relative cell references.
 *
 * @since 1.6.0
 */
public final class ConditionalFormats {

    private ConditionalFormats() {}

    // ─── greaterThan ───────────────────────────────────────────────

    public static FormatCondition greaterThan(final Number value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public static FormatCondition greaterThan(final LocalDate value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public static FormatCondition greaterThan(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public static FormatCondition greaterThan(final Date value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public static FormatCondition greaterThan(final Calendar value) {
        return _cellIs(ComparisonOperator.GT, _f(value));
    }

    public static FormatCondition greaterThan(final Formula formula) {
        Objects.requireNonNull(formula, "formula");
        return _cellIs(ComparisonOperator.GT, formula);
    }

    // ─── greaterThanOrEqual ────────────────────────────────────────

    public static FormatCondition greaterThanOrEqual(final Number value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public static FormatCondition greaterThanOrEqual(final LocalDate value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public static FormatCondition greaterThanOrEqual(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public static FormatCondition greaterThanOrEqual(final Date value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public static FormatCondition greaterThanOrEqual(final Calendar value) {
        return _cellIs(ComparisonOperator.GE, _f(value));
    }

    public static FormatCondition greaterThanOrEqual(final Formula formula) {
        Objects.requireNonNull(formula, "formula");
        return _cellIs(ComparisonOperator.GE, formula);
    }

    // ─── lessThan ──────────────────────────────────────────────────

    public static FormatCondition lessThan(final Number value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public static FormatCondition lessThan(final LocalDate value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public static FormatCondition lessThan(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public static FormatCondition lessThan(final Date value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public static FormatCondition lessThan(final Calendar value) {
        return _cellIs(ComparisonOperator.LT, _f(value));
    }

    public static FormatCondition lessThan(final Formula formula) {
        Objects.requireNonNull(formula, "formula");
        return _cellIs(ComparisonOperator.LT, formula);
    }

    // ─── lessThanOrEqual ───────────────────────────────────────────

    public static FormatCondition lessThanOrEqual(final Number value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public static FormatCondition lessThanOrEqual(final LocalDate value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public static FormatCondition lessThanOrEqual(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public static FormatCondition lessThanOrEqual(final Date value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public static FormatCondition lessThanOrEqual(final Calendar value) {
        return _cellIs(ComparisonOperator.LE, _f(value));
    }

    public static FormatCondition lessThanOrEqual(final Formula formula) {
        Objects.requireNonNull(formula, "formula");
        return _cellIs(ComparisonOperator.LE, formula);
    }

    // ─── equalTo ───────────────────────────────────────────────────

    public static FormatCondition equalTo(final Number value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public static FormatCondition equalTo(final boolean value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public static FormatCondition equalTo(final String value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public static FormatCondition equalTo(final LocalDate value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public static FormatCondition equalTo(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public static FormatCondition equalTo(final Date value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public static FormatCondition equalTo(final Calendar value) {
        return _cellIs(ComparisonOperator.EQUAL, _f(value));
    }

    public static FormatCondition equalTo(final Formula formula) {
        Objects.requireNonNull(formula, "formula");
        return _cellIs(ComparisonOperator.EQUAL, formula);
    }

    // ─── notEqualTo ────────────────────────────────────────────────

    public static FormatCondition notEqualTo(final Number value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public static FormatCondition notEqualTo(final boolean value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public static FormatCondition notEqualTo(final String value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public static FormatCondition notEqualTo(final LocalDate value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public static FormatCondition notEqualTo(final LocalDateTime value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public static FormatCondition notEqualTo(final Date value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public static FormatCondition notEqualTo(final Calendar value) {
        return _cellIs(ComparisonOperator.NOT_EQUAL, _f(value));
    }

    public static FormatCondition notEqualTo(final Formula formula) {
        Objects.requireNonNull(formula, "formula");
        return _cellIs(ComparisonOperator.NOT_EQUAL, formula);
    }

    // ─── between ───────────────────────────────────────────────────

    public static FormatCondition between(final Number low, final Number high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition between(final LocalDate low, final LocalDate high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition between(final LocalDateTime low, final LocalDateTime high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition between(final Date low, final Date high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition between(final Calendar low, final Calendar high) {
        return _cellIs(ComparisonOperator.BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition between(final Formula low, final Formula high) {
        Objects.requireNonNull(low, "low");
        Objects.requireNonNull(high, "high");
        return _cellIs(ComparisonOperator.BETWEEN, low, high);
    }

    // ─── notBetween ────────────────────────────────────────────────

    public static FormatCondition notBetween(final Number low, final Number high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition notBetween(final LocalDate low, final LocalDate high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition notBetween(final LocalDateTime low, final LocalDateTime high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition notBetween(final Date low, final Date high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition notBetween(final Calendar low, final Calendar high) {
        return _cellIs(ComparisonOperator.NOT_BETWEEN, _f(low), _f(high));
    }

    public static FormatCondition notBetween(final Formula low, final Formula high) {
        Objects.requireNonNull(low, "low");
        Objects.requireNonNull(high, "high");
        return _cellIs(ComparisonOperator.NOT_BETWEEN, low, high);
    }

    // ─── expression ────────────────────────────────────────────────

    /**
     * Arbitrary boolean Excel formula evaluated per cell in the conditional formatting
     * range. Maps to OOXML {@code <cfRule type="expression">}. The formula must return
     * TRUE or FALSE.
     *
     * <pre>{@code
     * grid.conditionalFormatting("price",
     *         expression("=AND($A1>0, $B1<100)").style("warn"));
     * }</pre>
     */
    public static FormatCondition expression(final String formula) {
        if (formula == null || formula.isEmpty()) {
            throw new IllegalArgumentException("Expression formula must not be empty");
        }
        return new ExpressionCondition(formula);
    }

    // ─── colorScale ────────────────────────────────────────────────

    /**
     * 3-color color scale with Excel's default thresholds ({@code MIN / PERCENTILE 50 / MAX})
     * and the standard red → yellow → green color ramp.
     */
    public static ConditionalFormatRule colorScale() {
        return new ColorScaleRule();
    }

    /**
     * 3-color color scale with explicit {@code NUMBER} thresholds and the standard
     * red → yellow → green color ramp.
     */
    public static ConditionalFormatRule colorScale(final double min, final double mid, final double max) {
        return new ColorScaleRule(min, mid, max);
    }

    // ─── operand factories ─────────────────────────────────────────

    /**
     * Raw Excel formula passthrough operand. The text is emitted verbatim into the OOXML
     * {@code <formula>} element. The library does not validate or parse it.
     *
     * <pre>{@code
     * grid.conditionalFormatting("price",
     *         greaterThan(formula("$D$1")).style("warn"),
     *         greaterThan(formula("AVERAGE($B$2:$B$100)")).style("aboveAvg"));
     * }</pre>
     */
    public static Formula formula(final String text) {
        if (text == null || text.isEmpty()) {
            throw new IllegalArgumentException("Formula text must not be empty");
        }
        return new Formula(Formula.Kind.OF, text);
    }

    /**
     * Schema-aware row-relative reference to another column in the data grid.
     * Resolved to {@code $<colLetter><dataStartRow>} at apply time, so Excel
     * auto-shifts the row for each cell in the conditional formatting range.
     *
     * <pre>{@code
     * grid.conditionalFormatting("price",
     *         greaterThan(columnRef("minPrice")).style("warn"));
     * }</pre>
     */
    public static Formula columnRef(final String name) {
        if (name == null || name.isEmpty()) {
            throw new IllegalArgumentException("Column name must not be empty");
        }
        return new Formula(Formula.Kind.COLUMN, name);
    }

    // ─── private helpers ───────────────────────────────────────────

    private static FormatCondition _cellIs(final byte operator, final Object operand) {
        return new CellIsCondition(operator, operand, null);
    }

    private static FormatCondition _cellIs(final byte operator, final Object low, final Object high) {
        return new CellIsCondition(operator, low, high);
    }

    private static String _f(final Number value) {
        Objects.requireNonNull(value, "value");
        return value.toString();
    }

    private static String _f(final boolean value) {
        return value ? "TRUE" : "FALSE";
    }

    private static String _f(final String value) {
        Objects.requireNonNull(value, "value");
        // Excel string literal: surrounded by quotes; internal quotes doubled
        return "\"" + value.replace("\"", "\"\"") + "\"";
    }

    private static String _f(final LocalDate value) {
        Objects.requireNonNull(value, "value");
        return "DATE(" + value.getYear() + ","
                + value.getMonthValue() + ","
                + value.getDayOfMonth() + ")";
    }

    private static String _f(final LocalDateTime value) {
        Objects.requireNonNull(value, "value");
        return _f(value.toLocalDate())
                + "+TIME(" + value.getHour() + ","
                + value.getMinute() + ","
                + value.getSecond() + ")";
    }

    private static String _f(final Date value) {
        Objects.requireNonNull(value, "value");
        return _f(LocalDateTime.ofInstant(value.toInstant(), ZoneId.systemDefault()));
    }

    private static String _f(final Calendar value) {
        Objects.requireNonNull(value, "value");
        return _f(LocalDateTime.ofInstant(value.toInstant(), value.getTimeZone().toZoneId()));
    }
}
