package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.w3c.dom.Document;

final class OpcXmlHelper {

    private static final DocumentBuilderFactory DBF = DocumentBuilderFactory.newInstance();
    static {
        DBF.setNamespaceAware(true);
    }

    static final String NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    static Document parsePart(final OPCPackage pkg, final String partName) throws Exception {
        try (InputStream is = pkg.getPart(
                PackagingURIHelper.createPartName(partName)).getInputStream()) {
            return DBF.newDocumentBuilder().parse(is);
        }
    }

    private OpcXmlHelper() {}
}
