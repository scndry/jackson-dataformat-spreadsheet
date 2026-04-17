package io.github.scndry.jackson.dataformat.spreadsheet.poi;

import java.lang.reflect.Method;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Compatibility shim for POI API differences across versions.
 * Uses reflection to call methods that may not exist in older
 * POI releases, falling back to safe defaults.
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
