package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import lombok.extern.slf4j.Slf4j;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetWriter;

/**
 * Streaming SpreadsheetML {@link SheetWriter} implementation for XLSX output.
 * Builds a temporary {@link XSSFWorkbook} skeleton for package compatibility, then rewrites worksheet and
 * sharedStrings parts directly while copying the remaining package entries from the skeleton.
 *
 * @see io.github.scndry.jackson.dataformat.spreadsheet.poi.ss.POISheetWriter
 */
@Slf4j
public final class SSMLSheetWriter implements SheetWriter {

    private static final String XML_DECL =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";

    private static final int BUFFER_SIZE = 512 * 1024;
    private static final int FLUSH_THRESHOLD = 1024;

    private final String _entrySheet;
    private final String _entrySst;

    private final ZipOutputStream _zip;
    private final StringBuilder _sb = new StringBuilder(BUFFER_SIZE);

    private SpreadsheetSchema _schema;
    private CellAddress _reference;
    private int _currentRow = -1;
    private boolean _currentRowOpen;

    // SharedStrings
    private final SharedStringsStore _sst;

    // Styles — resolved from POI XSSFWorkbook
    private XSSFWorkbook _wb;
    private File _skeleton;
    private int[] _columnStyleIndex;
    private int[] _headerColumnStyleIndex;
    private final List<MergeRange> _mergeRanges = new ArrayList<>();
    private XSSFSheet _sheet;
    private boolean _sheetDataStarted;
    private String _sheetXmlPrefix;
    private String _sheetXmlSuffix;

    public SSMLSheetWriter(final OutputStream os, final Sheet sheet) {
        this(os, sheet, false, false);
    }

    public SSMLSheetWriter(
            final OutputStream os,
            final Sheet sheet,
            final boolean fileBackedSharedStrings,
            final boolean encryptFileBackedStore) {
        _zip = new ZipOutputStream(os);
        _zip.setLevel(4);
        _sheet = (XSSFSheet) sheet;
        _sst = _createSharedStringsStore(fileBackedSharedStrings, encryptFileBackedStore);
        final PackagePart sheetPackagePart = _sheet.getPackagePart();
        final XSSFWorkbook workbook = _sheet.getWorkbook();
        final PackagePart sharedStringsPackagePart = workbook.getSharedStringSource().getPackagePart();

        _entrySheet = _entryName(sheetPackagePart); // xl/worksheets/sheetN.xml
        _entrySst = _entryName(sharedStringsPackagePart); // xl/sharedStrings.xml
    }

    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        return SpreadsheetVersion.EXCEL2007;
    }

    @Override
    public void setSchema(final SpreadsheetSchema schema) {
        _schema = schema;
        try {
            _wb = _sheet.getWorkbook();
            final Styles styles = _schema.buildStyles(_wb);
            _schema.applyFreezePane(_sheet);
            _schema.applyConditionalFormattings(_sheet, styles, -1);
            _writeSkeletonWorkbook();
            _splitSkeletonSheetXml();
            _columnStyleIndex = _resolveColumnStyleIndices(styles, false);
            _headerColumnStyleIndex = _resolveColumnStyleIndices(styles, true);
            _wb.close();
            _wb = null;
            _sheet = null;
            _startSheetXmlEntry();
            _appendFixedColumns();
            _startSheetData();
        } catch (IOException e) {
            throw new IllegalStateException("Failed to build skeleton", e);
        }
    }

    @Override
    public void setReference(final CellAddress reference) {
        _reference = reference;
        try {
            final int row = reference.getRow();
            if (row != _currentRow) {
                _closeCurrentRowIfOpen();
                _currentRow = row;
                _currentRowOpen = false;
            }
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    private void _ensureRowOpen() throws IOException {
        if (!_currentRowOpen) {
            _append("<row r=\"").append(_currentRow + 1).append("\">");
            _currentRowOpen = true;
        }
    }

    @Override
    public void writeHeaders() {
        final int row = _schema.getOriginRow();
        for (final Column column : _schema) {
            final int col = _schema.columnIndexOf(column);
            setReference(new CellAddress(row, col));
            writeString(column.getName());
        }
    }

    @Override
    public void writeNumeric(final double value) {
        try {
            _appendCellStart("n").append(value);
            _appendCellEnd();
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    @Override
    public void writeString(final String value) {
        try {
            final int index = _cacheString(value);
            _appendCellStart("s").append(index);
            _appendCellEnd();
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    @Override
    public void writeBoolean(final boolean value) {
        try {
            _appendCellStart("b").append(value ? '1' : '0');
            _appendCellEnd();
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    @Override
    public void writeBlank() {
        try {
            _ensureRowOpen();
            _append("<c r=\"").append(_cellRef())
                    .append("\" s=\"").append(_resolveStyleIndex())
                    .append("\"/>");
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    @Override
    public void adjustColumnWidth() {
        // no-op
    }

    @Override
    public void mergeScopedColumns(final ColumnPointer pointer, final int row, final int size) {
        if (size <= 1) return;
        final List<Column> columns = _schema.getColumns(pointer);
        for (final Column column : columns) {
            if (!column.isMerge()) continue;
            if (pointer.relativize(column.getPointer()).contains(ColumnPointer.array())) continue;
            final int col = _schema.columnIndexOf(column);
            if (col < 0) continue;
            _mergeRanges.add(new MergeRange(row, row + size - 1, col));
        }
    }

    @Override
    public void write() throws IOException {
        _finishSheetXmlEntry();
        _writeSharedStringsEntry();
        _copySkeletonEntries();
        _zip.finish();
    }

    @Override
    public void close() throws IOException {
        IOException failure = null;
        try {
            _zip.close();
        } catch (IOException e) {
            failure = e;
        }
        if (_wb != null) {
            try {
                _wb.close();
            } catch (IOException e) {
                failure = _mergeFailure(failure, e);
            }
        }
        try {
            _sst.close();
        } catch (IOException e) {
            failure = _mergeFailure(failure, e);
        }
        if (_skeleton != null && !_skeleton.delete() && log.isDebugEnabled()) {
            log.debug("Failed to delete skeleton temp file: {}", _skeleton);
        }
        if (failure != null) {
            throw failure;
        }
    }

    @Override
    public boolean isDate1904() {
        return false;
    }

    // ----------------------------------------------------------------
    // Skeleton copy
    // ----------------------------------------------------------------

    private void _copySkeletonEntries() throws IOException {
        final byte[] buf = new byte[8192];
        try (ZipInputStream zin = new ZipInputStream(new FileInputStream(_skeleton))) {
            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                if (_shouldSkipEntry(entry.getName())) {
                    zin.closeEntry();
                    continue;
                }
                _zip.putNextEntry(new ZipEntry(entry.getName()));
                int len;
                while ((len = zin.read(buf)) > 0) {
                    _zip.write(buf, 0, len);
                }
                _zip.closeEntry();
                zin.closeEntry();
            }
        }
    }

    // ----------------------------------------------------------------
    // StringBuilder streaming
    // ----------------------------------------------------------------

    private StringBuilder _append(final String s) throws IOException {
        _sb.append(s);
        _checkFlush();
        return _sb;
    }

    private StringBuilder _appendCellStart(final String type) throws IOException {
        _ensureRowOpen();
        return _append("<c r=\"").append(_cellRef())
                .append("\" s=\"").append(_resolveStyleIndex())
                .append("\" t=\"").append(type).append("\"><v>");
    }

    private void _appendCellEnd() throws IOException {
        _append("</v></c>");
    }

    private void _checkFlush() throws IOException {
        if (_sb.capacity() - _sb.length() < FLUSH_THRESHOLD) {
            _flush();
        }
    }

    private void _flush() throws IOException {
        _zip.write(_sb.toString().getBytes(StandardCharsets.UTF_8));
        _sb.setLength(0);
    }

    private String _cellRef() {
        return CellReference.convertNumToColString(_reference.getColumn())
            + (_reference.getRow() + 1);
    }

    private static String _cellRef(final int row, final int col) {
        return CellReference.convertNumToColString(col) + (row + 1);
    }

    private int _resolveStyleIndex() {
        if (_columnStyleIndex == null) return 0;
        final int idx = _reference.getColumn() - _schema.getOriginColumn();
        if (idx < 0 || idx >= _columnStyleIndex.length) return 0;
        if (_reference.getRow() < _schema.getDataRow()) {
            if (_headerColumnStyleIndex == null || idx >= _headerColumnStyleIndex.length) {
                return _columnStyleIndex[idx];
            }
            return _headerColumnStyleIndex[idx];
        }
        return _columnStyleIndex[idx];
    }

    private int _cacheString(final String value) {
        return _sst.getOrAdd(value);
    }

    private SharedStringsStore _createSharedStringsStore(
            final boolean fileBackedSharedStrings,
            final boolean encryptFileBackedStore) {
        try {
            return fileBackedSharedStrings
                    ? new FileBackedSharedStringsStore(encryptFileBackedStore)
                    : new InMemorySharedStringsStore(256);
        } catch (IOException e) {
            throw new IllegalStateException("Failed to initialize shared string store", e);
        }
    }

    private static String _entryName(final PackagePart part) {
        return part.getPartName().getName().substring(1);
    }

    private void _startSheetXmlEntry() throws IOException {
        _zip.putNextEntry(new ZipEntry(_entrySheet));
        _append(_sheetXmlPrefix);
    }

    private void _startSheetData() throws IOException {
        if (_sheetDataStarted) {
            return;
        }
        _append("<sheetData>");
        _sheetDataStarted = true;
    }

    private void _finishSheetXmlEntry() throws IOException {
        _startSheetData();
        _closeCurrentRowIfOpen();
        _append("</sheetData>");
        _appendMergeCellsIntoSuffix();
        _flush();
        _zip.closeEntry();
    }

    private void _closeCurrentRowIfOpen() throws IOException {
        if (_currentRowOpen) {
            _append("</row>");
            _currentRowOpen = false;
        }
    }

    private void _writeSharedStringsEntry() throws IOException {
        _zip.putNextEntry(new ZipEntry(_entrySst));
        _writeSharedStrings();
        _zip.closeEntry();
    }

    private void _writeSkeletonWorkbook() throws IOException {
        _skeleton = TempFile.createTempFile("jackson-spreadsheet-skeleton-", ".xlsx");
        try (FileOutputStream fos = new FileOutputStream(_skeleton)) {
            _wb.write(fos);
        }
    }

    private int[] _resolveColumnStyleIndices(final Styles styles, final boolean header) {
        final List<Column> columns = new ArrayList<>();
        for (final Column column : _schema) {
            columns.add(column);
        }

        final int[] columnStyleIndex = new int[columns.size()];
        for (int i = 0; i < columns.size(); i++) {
            final CellStyle cs = header
                    ? styles.getHeaderStyle(columns.get(i))
                    : styles.getStyle(columns.get(i));
            if (cs != null) {
                columnStyleIndex[i] = cs.getIndex();
            }
        }
        return columnStyleIndex;
    }

    private boolean _shouldSkipEntry(final String entryName) {
        return _entrySheet.equals(entryName) || _entrySst.equals(entryName);
    }

    private void _splitSkeletonSheetXml() throws IOException {
        try (ZipInputStream zin = new ZipInputStream(Files.newInputStream(_skeleton.toPath()))) {
            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                if (_entrySheet.equals(entry.getName())) {
                    final ByteArrayOutputStream bos = new ByteArrayOutputStream();
                    final byte[] buf = new byte[8192];
                    int n;
                    while ((n = zin.read(buf)) != -1) bos.write(buf, 0, n);
                    final String sheetXml = bos.toString("UTF-8");
                    final int sdStart = sheetXml.indexOf("<sheetData");
                    if (sdStart < 0) {
                        throw new IOException("Skeleton sheet XML missing <sheetData> element");
                    }
                    _sheetXmlPrefix = sheetXml.substring(0, sdStart);
                    final int sdEnd = sheetXml.indexOf("</sheetData>");
                    if (sdEnd >= 0) {
                        _sheetXmlSuffix = sheetXml.substring(sdEnd + "</sheetData>".length());
                    } else {
                        final int selfClose = sheetXml.indexOf("/>", sdStart);
                        if (selfClose < 0) {
                            throw new IOException("Malformed skeleton sheet XML: unclosed <sheetData> element");
                        }
                        _sheetXmlSuffix = sheetXml.substring(selfClose + 2);
                    }
                    return;
                }
                zin.closeEntry();
            }
        }
        throw new IOException("Sheet entry '" + _entrySheet + "' not found in skeleton");
    }

    private void _appendMergeCellsIntoSuffix() throws IOException {
        if (_mergeRanges.isEmpty()) {
            _append(_sheetXmlSuffix);
            return;
        }
        // ECMA-376 order: autoFilter → mergeCells → conditionalFormatting
        // Find insertion point in suffix for mergeCells
        final String mergeCellsXml = _buildMergeCellsXml();
        final int insertPos = _findMergeCellsInsertPos(_sheetXmlSuffix);
        _append(_sheetXmlSuffix.substring(0, insertPos));
        _append(mergeCellsXml);
        _append(_sheetXmlSuffix.substring(insertPos));
    }

    private String _buildMergeCellsXml() {
        final StringBuilder sb = new StringBuilder();
        sb.append("<mergeCells count=\"").append(_mergeRanges.size()).append("\">");
        for (final MergeRange range : _mergeRanges) {
            sb.append("<mergeCell ref=\"")
                    .append(_cellRef(range._firstRow, range._column))
                    .append(":")
                    .append(_cellRef(range._lastRow, range._column))
                    .append("\"/>");
        }
        sb.append("</mergeCells>");
        return sb.toString();
    }

    // ECMA-376 CT_Worksheet elements that follow mergeCells in xsd:sequence order.
    // Used to locate the correct insertion point for <mergeCells> in the skeleton suffix.
    private static final String[] POST_MERGE_CELLS_ELEMENTS = {
            "<phoneticPr", "<conditionalFormatting", "<dataValidations",
            "<hyperlinks", "<printOptions", "<pageMargins", "<pageSetup",
            "<headerFooter", "<rowBreaks", "<colBreaks", "<drawing",
            "<tableParts", "<extLst", "</worksheet"
    };

    private static int _findMergeCellsInsertPos(final String suffix) {
        for (final String tag : POST_MERGE_CELLS_ELEMENTS) {
            final int pos = suffix.indexOf(tag);
            if (pos >= 0) return pos;
        }
        return suffix.length();
    }

    private static IOException _mergeFailure(final IOException failure, final IOException next) {
        if (failure == null) {
            return next;
        }
        failure.addSuppressed(next);
        return failure;
    }

    private void _writeSharedStrings() throws IOException {
        _sb.setLength(0);
        _append(XML_DECL);
        _append("<sst xmlns=\"").append(XSSFRelation.NS_SPREADSHEETML)
            .append("\" count=\"").append(_sst.size())
            .append("\" uniqueCount=\"").append(_sst.size()).append("\">");
        for (int i = 0; i < _sst.size(); i++) {
            _appendSharedStringItem(i);
            _checkFlush();
        }
        _append("</sst>");
        _flush();
    }

    private void _appendSharedStringItem(final int index) throws IOException {
        final int len = _sst.length(index);
        _sb.append("<si><t");
        if (len == 0
                || Character.isWhitespace(_sst.charAt(index, 0))
                || Character.isWhitespace(_sst.charAt(index, len - 1))) {
            _sb.append(" xml:space=\"preserve\"");
        }
        _sb.append(">");
        _appendEscapedSharedString(index, len);
        _sb.append("</t></si>");
    }

    private void _appendEscapedSharedString(final int index, final int len) {
        for (int i = 0; i < len; i++) {
            final char c = _sst.charAt(index, i);
            switch (c) {
                case '<':
                    _sb.append("&lt;");
                    break;
                case '>':
                    _sb.append("&gt;");
                    break;
                case '&':
                    _sb.append("&amp;");
                    break;
                case '\'':
                    _sb.append("&apos;");
                    break;
                case '"':
                    _sb.append("&quot;");
                    break;
                default:
                    if (c >= 0x20 || c == 0x9 || c == 0xA || c == 0xD) {
                        _sb.append(c);
                    }
            }
        }
    }

    private void _appendFixedColumns() throws IOException {
        if (_schema == null) {
            return;
        }
        boolean hasFixedWidth = false;
        for (final Column column : _schema) {
            if (column != null && column.getValue().getWidth() > 0) {
                hasFixedWidth = true;
                break;
            }
        }
        if (!hasFixedWidth) {
            return;
        }
        _append("<cols>");
        for (final Column column : _schema) {
            if (column == null) continue;
            final int width = Math.min(column.getValue().getWidth(), 255);
            if (width <= 0) continue;
            final int col = _schema.columnIndexOf(column) + 1;
            _append("<col min=\"").append(col)
                    .append("\" max=\"").append(col)
                    .append("\" width=\"").append(width)
                    .append("\" customWidth=\"1\"/>");
        }
        _append("</cols>");
    }

    private static final class MergeRange {
        final int _firstRow;
        final int _lastRow;
        final int _column;

        MergeRange(final int firstRow, final int lastRow, final int column) {
            _firstRow = firstRow;
            _lastRow = lastRow;
            _column = column;
        }
    }

}
