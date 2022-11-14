package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xssf.usermodel.XSSFRelation;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;

@Slf4j
final class XSSFCorePart {

    private final PackagePart _part;

    XSSFCorePart(final PackagePart corePart) {
        _part = corePart;
    }

    public InputStream getInputStream() {
        try {
            return _part.getInputStream();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    public PackagePart getRelatedPart(final XSSFRelation rel) {
        return getRelatedPart(getRelationshipByType(rel));
    }

    public PackagePart getRelatedPart(final PackageRelationship rel) {
        if (!_part.isRelationshipExists(rel)) return null;
        try {
            return _part.getRelatedPart(rel);
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(e);
        }
    }

    public PackageRelationship getRelationshipByType(final XSSFRelation rel) {
        return getRelationshipsByType(rel).getRelationship(0);
    }

    public PackageRelationshipCollection getRelationshipsByType(final XSSFRelation rel) {
        try {
            return _part.getRelationshipsByType(rel.getRelation());
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(e);
        }
    }

    public OPCPackage getPackage() {
        return _part.getPackage();
    }
}
