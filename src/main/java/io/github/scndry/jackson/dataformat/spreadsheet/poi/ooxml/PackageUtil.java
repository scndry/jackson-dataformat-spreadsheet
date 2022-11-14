package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;

@Slf4j
public final class PackageUtil {

    private PackageUtil() {
    }

    public static boolean isOOXML(final File source) {
        try {
            return FileMagic.valueOf(source) == FileMagic.OOXML;
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    public static boolean isOOXML(final InputStream source) {
        try {
            return FileMagic.valueOf(FileMagic.prepareToCheckMagic(source)) == FileMagic.OOXML;
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    public static OPCPackage open(final File source) {
        try {
            return OPCPackage.open(source, PackageAccess.READ);
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(e);
        }
    }

    public static OPCPackage open(final InputStream source) {
        try {
            return OPCPackage.open(source);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(e);
        }
    }

    public static PackagePart extractCorePart(final File source) {
        return _extractCorePart(open(source));
    }

    public static PackagePart extractCorePart(final InputStream source) {
        return _extractCorePart(open(source));
    }

    private static PackagePart _extractCorePart(final OPCPackage pack) {
        final PackageRelationship rel = _extractCoreRelationship(pack);
        if (rel == null) {
            throw new IllegalArgumentException("No core document relationship");
        }
        final PackagePartName name;
        try {
            name = PackagingURIHelper.createPartName(rel.getTargetURI());
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(e);
        }
        return pack.getPart(name);
    }

    private static PackageRelationship _extractCoreRelationship(final OPCPackage pack) {
        final PackageRelationship rel = pack.getRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT).getRelationship(0);
        if (rel != null) return rel;
        return pack.getRelationshipsByType(PackageRelationshipTypes.STRICT_CORE_DOCUMENT).getRelationship(0);
    }
}
