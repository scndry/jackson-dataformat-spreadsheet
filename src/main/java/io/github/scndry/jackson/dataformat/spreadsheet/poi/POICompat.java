package io.github.scndry.jackson.dataformat.spreadsheet.poi;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.EnumSet;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

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
     * Tracked temp files awaiting cleanup. The shutdown hook deletes any
     * entries still present at JVM exit. Normal release paths
     * ({@link #releaseTempFile}) remove their entry, so steady-state usage
     * does not accumulate — avoiding the unbounded growth of
     * {@link File#deleteOnExit()} (which has no unregister API).
     */
    private static final Set<Path> TRACKED = ConcurrentHashMap.newKeySet();

    static {
        Runtime.getRuntime().addShutdownHook(new Thread(POICompat::_cleanupTrackedOnShutdown,
                "jackson-spreadsheet-temp-cleanup"));
    }

    /**
     * Creates a temporary file in POI's temp directory under a dedicated
     * {@code jackson-spreadsheet} subdirectory, with owner-only permissions
     * on POSIX systems. Falls back to default permissions on Windows.
     * <p>
     * The returned path is tracked for shutdown-time cleanup. Callers must
     * invoke {@link #releaseTempFile(Path)} when the file is no longer
     * needed so the tracker entry is removed.
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
        TRACKED.add(path);
        return path;
    }

    /**
     * Deletes the temp file and removes it from the shutdown tracker.
     * Safe to call on a path that was never tracked or already released.
     * <p>
     * Delete is attempted before the tracker is updated: if delete fails,
     * the path stays in the tracker so the shutdown hook can retry — useful
     * for transient failures (full disk, Windows file lock) that may resolve
     * by JVM exit time.
     */
    public static void releaseTempFile(final Path path) throws IOException {
        Files.deleteIfExists(path);
        TRACKED.remove(path);
    }

    private static void _cleanupTrackedOnShutdown() {
        for (final Path p : TRACKED) {
            try {
                Files.deleteIfExists(p);
            } catch (Exception ignored) {
            }
        }
        TRACKED.clear();
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
