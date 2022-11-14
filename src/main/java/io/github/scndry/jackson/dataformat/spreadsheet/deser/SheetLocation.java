package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import com.fasterxml.jackson.core.JsonLocation;
import com.fasterxml.jackson.core.io.ContentReference;
import lombok.EqualsAndHashCode;
import org.apache.poi.ss.util.CellReference;

@EqualsAndHashCode(callSuper = true)
public final class SheetLocation extends JsonLocation {

    public static final String UNKNOWN = "UNKNOWN";
    private final int _row;
    private final int _column;

    public SheetLocation(final ContentReference contentRef, final int row, final int column) {
        super(contentRef, -1, -1, -1);
        _row = row;
        _column = column;
    }

    public int getRow() {
        return _row;
    }

    public int getColumn() {
        return _column;
    }

    @Override
    public String sourceDescription() {
        final Object content = _contentReference.getRawContent();
        if (content instanceof SheetInput) {
            return content.toString();
        }
        return super.sourceDescription();
    }

    @Override
    public StringBuilder appendOffsetDescription(final StringBuilder sb) {
        sb.append("row: ");
        if (_row >= 0) {
            sb.append(_row);
        } else {
            sb.append(UNKNOWN);
        }
        sb.append(", column: ");
        if (_column >= 0) {
            sb.append(_column);
        } else {
            sb.append(UNKNOWN);
        }
        sb.append(", address: ");
        if (_row >= 0 && _column >= 0) {
            sb.append(CellReference.convertNumToColString(_column));
            sb.append(_row + 1);
        } else {
            sb.append(UNKNOWN);
        }
        return sb;
    }
}
