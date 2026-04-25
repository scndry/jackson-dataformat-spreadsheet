package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.NoSuchElementException;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.util.CellAddress;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.CellValue;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetToken;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.CTCell;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.STCellFormulaType;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.SpreadsheetML;

/**
 * Streaming SpreadsheetML {@link SheetReader} implementation.
 * Parses worksheet XML via {@link XmlElementReader} and resolves shared strings without POI's User Model.
 *
 * @see io.github.scndry.jackson.dataformat.spreadsheet.poi.ss.POISheetReader
 * @see XmlElementReader
 */
@Slf4j
public final class SSMLSheetReader implements SheetReader {

    private static final Matcher START_SHEET_DATA = Matcher.startElement(SpreadsheetML.SHEET_DATA);
    private static final Matcher START_ROW = Matcher.startElement(SpreadsheetML.ROW);
    private static final Matcher END_ROW = Matcher.endElement(SpreadsheetML.ROW);
    private static final Matcher START_CELL = Matcher.startElement(SpreadsheetML.CELL);
    private static final Matcher END_SHEET_DATA = Matcher.endElement(SpreadsheetML.SHEET_DATA);

    private final SharedStringsLookup _strings;
    private final XmlElementReader _reader;
    private final SSMLWorkbook _workbook;
    private final PackagePart _sheet;
    private SheetToken _next;
    private CTCell _cell;
    private CellAddress _reference;
    private int _rowIndex = -1;
    private int _columnIndex = -1;

    public SSMLSheetReader(
            final PackagePart worksheetPart,
            final SSMLWorkbook workbook,
            final boolean fileBackedSharedStrings,
            final boolean encryptFileBacked) {
        _sheet = worksheetPart;
        _workbook = workbook;
        try {
            final PackagePart sharedStrings = _workbook.getSharedStringsPart();
            if (sharedStrings == null) {
                _strings = BlankSharedStringsLookup.INSTANCE;
            } else if (fileBackedSharedStrings) {
                _strings = new FileBackedSharedStringsLookup(sharedStrings, encryptFileBacked);
            } else {
                _strings = new InMemorySharedStringsLookup(sharedStrings);
            }
            _reader = new XmlElementReader(_sheet.getInputStream());
            _reader.nextUntil(START_SHEET_DATA);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        _next = SheetToken.SHEET_DATA_START;
    }

    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        return SpreadsheetVersion.EXCEL2007;
    }

    @Override
    public boolean isDate1904() {
        return _workbook.isDate1904();
    }

    @Override
    public CellAddress getReference() {
        return _reference;
    }

    @Override
    public CellValue getCellValue() {
        if (_cell == null) return null;
        final String value = _cell.getV();

        switch (_cell.getT()) {
            case BOOLEAN:
                return CellValue.valueOf("1".equals(value));
            case NUMBER:
                if (_cell.getFt() == null) {
                    return value == null ? CellValue.BLANK : new CellValue(
                            Double.parseDouble(value),
                            value);
                }
                if (_cell.getFt() == STCellFormulaType.SHARED) {
                    return new CellValue(Double.parseDouble(value));
                }
                throw new UnsupportedOperationException("Unexpected formula type: " +
                        _cell.getFt());
            case ERROR:
                return CellValue.getError(FormulaError.forString(value).getCode());
            case SHARED_STRING:
                return new CellValue(_strings.getItemAt(Integer.parseInt(value)));
            case FORMULA_STRING:
                return new CellValue(value);
            case INLINE_STRING:
                return _cell.getIs() != null ? new CellValue(_cell.getIs()) : new CellValue(value);
            default:
                throw new IllegalStateException();
        }
    }

    @Override
    public int getRow() {
        return _rowIndex;
    }

    @Override
    public int getColumn() {
        return _columnIndex;
    }

    @Override
    public boolean isClosed() {
        return _reader.isClosed();
    }

    @Override
    public void close() throws IOException {
        try {
            _strings.close();
        } catch (Exception e) {
            throw new IOException("Failed to close shared strings", e);
        }
        _workbook.close();
        try {
            _sheet.close();
        } catch (InvalidOperationException e) {
            // PackagePart.close() throws if the package is already closed by _workbook.close()
        }
        _reader.close();
    }

    @Override
    public boolean hasNext() {
        return _next != null;
    }

    @Override
    public SheetToken next() {
        if (_next == null) throw new NoSuchElementException();
        final SheetToken token = _next;
        switch (token) {
            case SHEET_DATA_START:
                _next = _matched(
                        START_ROW,
                        END_SHEET_DATA) ? SheetToken.ROW_START : SheetToken.SHEET_DATA_END;
                break;
            case ROW_START:
                _rowIndex = _reader.intAttribute(SpreadsheetML.ATTR_REF) - 1;
                _next = _matched(START_CELL, END_ROW) ? SheetToken.CELL_VALUE : SheetToken.ROW_END;
                break;
            case CELL_VALUE:
                _cell = _reader.collectCell();
                _reference = new CellAddress(_cell.getR());
                _columnIndex = _reference.getColumn();
                _next = _matched(START_CELL, END_ROW) ? SheetToken.CELL_VALUE : SheetToken.ROW_END;
                break;
            case ROW_END:
                _cell = null;
                _reference = null;
                _next = _matched(
                        START_ROW,
                        END_SHEET_DATA) ? SheetToken.ROW_START : SheetToken.SHEET_DATA_END;
                break;
            case SHEET_DATA_END:
                _next = null;
                break;
        }
        if (log.isTraceEnabled()) {
            if (token == SheetToken.CELL_VALUE) {
                log.trace("{} {} {}", token, getReference(), getCellValue());
            } else {
                log.trace("{}", token);
            }
        }
        return token;
    }

    private boolean _matched(final Matcher start, final Matcher end) {
        final Matcher hit = _reader.nextUntil(start, end);
        return hit != null && !hit.isEndElement();
    }
}
