package io.github.scndry.jackson.dataformat.spreadsheet.poi;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.EnumSet;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.TempFile;

/**
 * Compatibility shim for POI API differences across versions
 * and POI resource integration (temp file directory).
 */
public final class POICompat {

    private static final Method DATE_1904_METHOD;
    private static final Method STRICT_FORMAT_METHOD;

    static {
        Method d = null;
        try {
            d = Class.forName("org.apache.poi.ss.usermodel.Date1904Support")
                .getMethod("isDate1904");
        } catch (Exception ignored) {
        }
        DATE_1904_METHOD = d;

        Method s = null;
        try {
            s = OPCPackage.class.getMethod("isStrictOoxmlFormat");
        } catch (Exception ignored) {
        }
        STRICT_FORMAT_METHOD = s;
    }

    private POICompat() {
    }

    /** POI 4.1.1+ {@code Date1904Support.isDate1904()}. */
    public static boolean isDate1904(final Workbook workbook) {
        if (DATE_1904_METHOD == null) return false;
        if (!DATE_1904_METHOD.getDeclaringClass().isInstance(workbook)) return false;
        try {
            return (Boolean) DATE_1904_METHOD.invoke(workbook);
        } catch (Exception e) {
            return false;
        }
    }

    private static volatile File _tempDir;

    /**
     * Creates a temporary file in POI's temp directory under a dedicated
     * {@code jackson-spreadsheet} subdirectory, with owner-only permissions
     * on POSIX systems. Falls back to default permissions on Windows.
     */
    public static Path createSecureTempFile(final String prefix, final String suffix) throws IOException {
        final Path dir = tempDir().toPath();
        Path path;
        try {
            path = Files.createTempFile(dir, prefix, suffix,
                    PosixFilePermissions.asFileAttribute(
                            EnumSet.of(PosixFilePermission.OWNER_READ,
                                    PosixFilePermission.OWNER_WRITE)));
        } catch (UnsupportedOperationException e) {
            path = Files.createTempFile(dir, prefix, suffix);
        }
        path.toFile().deleteOnExit();
        return path;
    }

    /**
     * Returns a dedicated temp directory under POI's temp directory.
     * Created once and cached. Used by all library temp files.
     */
    public static File tempDir() throws IOException {
        File dir = _tempDir;
        if (dir != null && dir.isDirectory()) return dir;
        synchronized (POICompat.class) {
            dir = _tempDir;
            if (dir != null && dir.isDirectory()) return dir;
            dir = TempFile.createTempDirectory("jackson-spreadsheet");
            _tempDir = dir;
            return dir;
        }
    }

    /** POI 5.1.0+ {@code OPCPackage.isStrictOoxmlFormat()}. */
    public static boolean isStrictOoxmlFormat(
        final OPCPackage pkg) {
        if (STRICT_FORMAT_METHOD == null) return false;
        try {
            return (Boolean) STRICT_FORMAT_METHOD.invoke(pkg);
        } catch (Exception e) {
            return false;
        }
    }
}
