package io.github.scndry.jackson.dataformat.spreadsheet.schema.grid;

/**
 * Internal colorScale ({@code COLOR_SCALE}) 3-color rule.
 *
 * <p>Two flavors:
 * <ul>
 *   <li>{@code _useExcelDefaults = true}: Excel defaults — {@code MIN / PERCENTILE 50 / MAX}
 *       thresholds with the standard red / yellow / green color ramp.</li>
 *   <li>{@code _useExcelDefaults = false}: explicit {@code NUMBER} thresholds with the
 *       same default colors.</li>
 * </ul>
 *
 * <p>Color customization and threshold types other than {@code NUMBER} are deferred to
 * a later 1.6.x release.
 */
final class ColorScaleRule extends ConditionalFormatRule {

    final boolean _useExcelDefaults;
    final double _min;
    final double _mid;
    final double _max;

    /** Excel defaults: MIN / PERCENTILE 50 / MAX with standard color ramp. */
    ColorScaleRule() {
        _useExcelDefaults = true;
        _min = 0;
        _mid = 0;
        _max = 0;
    }

    /** Explicit NUMBER thresholds with default red/yellow/green color ramp. */
    ColorScaleRule(final double min, final double mid, final double max) {
        _useExcelDefaults = false;
        _min = min;
        _mid = mid;
        _max = max;
    }
}
