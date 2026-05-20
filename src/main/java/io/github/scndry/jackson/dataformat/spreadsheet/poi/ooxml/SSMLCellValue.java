package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import org.apache.poi.ss.usermodel.DataFormatter;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.CellValue;

/**
 * {@link CellValue} for SSML-mode NUMERIC cells that defers
 * {@link DataFormatter#formatRawCellContents} until {@link #getStringValue()}.
 *
 * @see SSMLSheetReader
 */
final class SSMLCellValue extends CellValue {

    private final int _styleIndex;
    private final StylesLookup _styles;
    private final DataFormatter _formatter;
    private String _lazyText;
    private boolean _computed;

    SSMLCellValue(final double numberValue, final String rawText,
                  final int styleIndex, final StylesLookup styles,
                  final DataFormatter formatter) {
        super(numberValue, rawText);
        _styleIndex = styleIndex;
        _styles = styles;
        _formatter = formatter;
    }

    @Override
    public String getStringValue() {
        if (!_computed) {
            _lazyText = _formattedString();
            _computed = true;
        }
        return _lazyText;
    }

    private String _formattedString() {
        final int numFmtId = _styles.getNumFmtId(_styleIndex);
        final String fmt = _styles.getFormatString(numFmtId);
        if (fmt == null) return super.getStringValue();
        return _formatter.formatRawCellContents(
                getNumberValue(),
                numFmtId,
                fmt);
    }
}
