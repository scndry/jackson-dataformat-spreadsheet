package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSst;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

final class LazySharedStrings implements SharedStrings {

    private static final Matcher START_SST = Matcher.startElementOf(CTSst.type);
    private static final Matcher START_RST = Matcher.startElementOf(CTRst.type);

    private final XmlElementReader _reader;
    private final int _count;
    private final int _uniqueCount;
    private final List<String> _strings;

    LazySharedStrings(final PackagePart part) throws IOException {
        _reader = new XmlElementReader(part.getInputStream());
        _reader.nextUntil(START_SST);
        final CTSst sst = _reader.current();
        _count = (int) sst.getUniqueCount();
        _uniqueCount = (int) sst.getUniqueCount();
        _strings = new ArrayList<>(_uniqueCount);
    }

    @Override
    public RichTextString getItemAt(final int idx) {
        while (idx >= _strings.size()) {
            _reader.nextUntil(START_RST);
            final CTRst stringItem = _reader.collect();
            final String text = stringItem.getT();
            if (text == null) {
                final StringBuilder sb = new StringBuilder();
                for (final CTRElt richTextRun : stringItem.getRArray()) {
                    sb.append(richTextRun.getT());
                }
                _strings.add(sb.toString());
            } else {
                _strings.add(text);
            }
        }
        return new XSSFRichTextString(_strings.get(idx));
    }

    @Override
    public int getCount() {
        return _count;
    }

    @Override
    public int getUniqueCount() {
        return _uniqueCount;
    }
}
