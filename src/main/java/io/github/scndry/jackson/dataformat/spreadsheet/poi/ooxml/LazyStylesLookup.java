package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.BuiltinFormats;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;

/**
 * {@link StylesLookup} that scans {@code <cellXfs>} and user-defined {@code <numFmts>}
 * via {@link XmlElementReader}; built-in number formats fall through to
 * {@link BuiltinFormats}. Parses on first lookup.
 */
final class LazyStylesLookup implements StylesLookup {

    private static final Matcher START_NUM_FMTS = Matcher.startElement(SpreadsheetML.NUM_FMTS);
    private static final Matcher START_NUM_FMT = Matcher.startElement(SpreadsheetML.NUM_FMT);
    private static final Matcher END_NUM_FMTS = Matcher.endElement(SpreadsheetML.NUM_FMTS);
    private static final Matcher START_CELL_XFS = Matcher.startElement(SpreadsheetML.CELL_XFS);
    private static final Matcher START_XF = Matcher.startElement(SpreadsheetML.XF);
    private static final Matcher END_CELL_XFS = Matcher.endElement(SpreadsheetML.CELL_XFS);
    private static final Matcher END_STYLE_SHEET = Matcher.endElement(SpreadsheetML.STYLE_SHEET);

    private final PackagePart _stylesPart;
    private int[] _xfNumFmtId;
    private Map<Integer, String> _userFormats;
    private boolean _initialized;

    LazyStylesLookup(final PackagePart stylesPart) {
        _stylesPart = stylesPart;
    }

    @Override
    public int getNumFmtId(final int xfIdx) {
        _init();
        if (xfIdx < 0 || xfIdx >= _xfNumFmtId.length) return 0;
        return _xfNumFmtId[xfIdx];
    }

    @Override
    public String getFormatString(final int numFmtId) {
        _init();
        final String custom = _userFormats.get(numFmtId);
        return custom != null ? custom : BuiltinFormats.getBuiltinFormat(numFmtId);
    }

    private void _init() {
        if (_initialized) return;
        final List<Integer> xfNumFmtIds = new ArrayList<>();
        final Map<Integer, String> userFormats = new HashMap<>();
        try (XmlElementReader reader = new XmlElementReader(_stylesPart.getInputStream())) {
            reader.navigateTo(SpreadsheetML.STYLE_SHEET);
            while (true) {
                final Matcher hit = reader.nextUntil(START_NUM_FMTS, START_CELL_XFS, END_STYLE_SHEET);
                if (hit == null || hit == END_STYLE_SHEET) break;
                if (hit == START_NUM_FMTS) {
                    _collectNumFmts(reader, userFormats);
                } else {
                    _collectCellXfs(reader, xfNumFmtIds);
                }
            }
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        _xfNumFmtId = new int[xfNumFmtIds.size()];
        for (int i = 0; i < _xfNumFmtId.length; i++) {
            _xfNumFmtId[i] = xfNumFmtIds.get(i);
        }
        _userFormats = userFormats;
        _initialized = true;
    }

    private static void _collectNumFmts(final XmlElementReader reader, final Map<Integer, String> out) {
        while (true) {
            final Matcher hit = reader.nextUntil(START_NUM_FMT, END_NUM_FMTS);
            if (hit == null || hit.isEndElement()) break;
            final int id = reader.intAttribute(SpreadsheetML.ATTR_NUM_FMT_ID);
            final String fmt = reader.attribute(SpreadsheetML.ATTR_FORMAT_CODE);
            if (fmt != null) out.put(id, fmt);
        }
    }

    private static void _collectCellXfs(final XmlElementReader reader, final List<Integer> out) {
        while (true) {
            final Matcher hit = reader.nextUntil(START_XF, END_CELL_XFS);
            if (hit == null || hit.isEndElement()) break;
            final String idAttr = reader.attribute(SpreadsheetML.ATTR_NUM_FMT_ID);
            out.add(idAttr == null ? 0 : Integer.parseInt(idAttr));
            reader.skipElement();
        }
    }
}
