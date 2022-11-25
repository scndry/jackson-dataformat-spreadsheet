package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.CellValue;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetToken;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.SchemaType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.NoSuchElementException;
import java.util.Objects;

@Slf4j
public final class SSMLSheetReader implements SheetReader {

    private static final Matcher START_SHEET_DATA = Matcher.startElementOf(CTSheetData.type);
    private static final Matcher END_SHEET_DATA = Matcher.endElementOf(CTSheetData.type);
    private static final Matcher START_ROW = Matcher.startElementOf(CTRow.type);
    private static final Matcher END_ROW = Matcher.endElementOf(CTRow.type);
    private static final Matcher START_CELL = Matcher.startElementOf(CTCell.type);

    private final SharedStrings _strings;
    private final XmlElementReader _reader;
    private final SSMLWorkbook _workbook;
    private final PackagePart _sheet;
    private SheetToken _next;
    private CTCell _cell;
    private CellAddress _reference;
    private int _rowIndex = -1;
    private int _columnIndex = -1;

    public SSMLSheetReader(final PackagePart worksheetPart, final SSMLWorkbook workbook) {
        _sheet = worksheetPart;
        _workbook = workbook;
        try {
            final PackagePart sharedStrings = _workbook.getSharedStringsPart();
            _strings = sharedStrings == null ? new BlankSharedStrings() : new LazySharedStrings(sharedStrings);
            _reader = new XmlElementReader(_sheet.getInputStream());
            final SchemaType type = _reader.getElementType();
            if (!type.equals(WorksheetDocument.type)) {
                throw new IllegalArgumentException("Unexpected package part: " + type);
            }
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        _reader.nextUntil(START_SHEET_DATA);
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
        final CTCellFormula formula = _cell.getF();
        final STCellType.Enum cellType = _cell.getT();
        switch (cellType.intValue()) {
            case STCellType.INT_B:
                return CellValue.valueOf(value.equals("1"));
            case STCellType.INT_N:
                if (formula == null) {
                    return value == null ? CellValue.BLANK : new CellValue(Double.parseDouble(value), value);
                }
                final STCellFormulaType.Enum formularType = formula.getT();
                if (formularType.equals(STCellFormulaType.SHARED)) {
                    return new CellValue(Double.parseDouble(value));
                }
                throw new UnsupportedOperationException("Unexpected cell formular type: " + formularType);
            case STCellType.INT_E:
                return CellValue.getError(FormulaError.forString(value).getCode());
            case STCellType.INT_S:
                return new CellValue(_strings.getItemAt(Integer.parseInt(value)).toString());
            case STCellType.INT_STR:
                return new CellValue(value);
            case STCellType.INT_INLINE_STR:
                final XSSFRichTextString text = _cell.isSetIs() ? new XSSFRichTextString(_cell.getIs()) : new XSSFRichTextString(value);
                return new CellValue(text.toString());
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
        _workbook.close();
        try {
            _sheet.close();
        } catch (InvalidOperationException e) { /* ignore */ }
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
                _next = _reader.nextUntil(START_ROW) == null ? SheetToken.SHEET_DATA_END : SheetToken.ROW_START;
                break;
            case ROW_START:
                final CTRow row = _reader.current();
                _rowIndex = (int) row.getR() - 1;
                _next = Objects.requireNonNull(_reader.nextUntil(START_CELL, END_ROW))
                        .isEndElement() ? SheetToken.ROW_END : SheetToken.CELL_VALUE;
                break;
            case CELL_VALUE:
                _cell = _reader.collect();
                _reference = new CellAddress(_cell.getR());
                _columnIndex = _reference.getColumn();
                _next = Objects.requireNonNull(_reader.nextUntil(START_CELL, END_ROW))
                        .isEndElement() ? SheetToken.ROW_END : SheetToken.CELL_VALUE;
                break;
            case ROW_END:
                _cell = null;
                _reference = null;
                _next = Objects.requireNonNull(_reader.nextUntil(START_ROW, END_SHEET_DATA))
                        .isEndElement() ? SheetToken.SHEET_DATA_END : SheetToken.ROW_START;
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

    static class BlankSharedStrings implements SharedStrings {

        @Override
        public RichTextString getItemAt(final int idx) {
            return new XSSFRichTextString("");
        }

        @Override
        public int getCount() {
            return 0;
        }

        @Override
        public int getUniqueCount() {
            return 0;
        }
    }
}
