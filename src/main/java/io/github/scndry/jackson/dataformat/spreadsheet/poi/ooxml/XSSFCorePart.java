package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xssf.usermodel.XSSFRelation;

/**
 * Wrapper around an OPC {@link PackagePart} representing the workbook core part.
 * Handles relationship resolution and Strict/Transitional namespace detection.
 *
 * @see PackageUtil
 */
@Slf4j
final class XSSFCorePart {

    private final PackagePart _part;
    private final boolean _strictFormat;

    XSSFCorePart(final PackagePart corePart) {
        _part = corePart;
        _strictFormat = corePart.getPackage().isStrictOoxmlFormat();
    }

    public InputStream getInputStream() {
        try {
            return _part.getInputStream();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    public PackagePart getRelatedPart(final XSSFRelation rel) {
        return getRelatedPart(getRelationship(rel));
    }

    public PackagePart getRelatedPart(final PackageRelationship rel) {
        if (!_part.isRelationshipExists(rel)) return null;
        try {
            return _part.getRelatedPart(rel);
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(e);
        }
    }

    public PackageRelationship getRelationship(final XSSFRelation rel) {
        return getRelationships(rel).getRelationship(0);
    }

    public PackageRelationshipCollection getRelationships(final XSSFRelation rel) {
        try {
            return _part.getRelationshipsByType(_relationshipType(rel));
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(e);
        }
    }

    public OPCPackage getPackage() {
        return _part.getPackage();
    }

    private String _relationshipType(final XSSFRelation rel) {
        if (_strictFormat) {
            return OoxmlNamespace.toStrict(rel.getRelation());
        }
        return rel.getRelation();
    }
}
