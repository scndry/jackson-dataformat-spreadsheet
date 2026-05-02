package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.poi.Version;
import org.apache.poi.openxml4j.opc.OPCPackage;

/**
 * Probes the runtime POI version for feature gates in tests.
 * The library targets POI 4.1.1+; some tests assert behavior that diverges
 * across POI versions and are guarded with the predicates here.
 */
public final class PoiVersionProbe {

    private PoiVersionProbe() {
    }

    /**
     * @return {@code true} if the runtime POI is 5.1.0 or later. Detected by probing
     * for {@code OPCPackage#isStrictOoxmlFormat()}, which was added in POI 5.1.0
     * (see {@code POICompat.STRICT_FORMAT_METHOD}).
     */
    public static boolean isPoi510OrLater() {
        try {
            OPCPackage.class.getMethod("isStrictOoxmlFormat");
            return true;
        } catch (NoSuchMethodException e) {
            return false;
        }
    }

    /**
     * @return {@code true} if the runtime POI is 5.2.3 or later. Determined by parsing
     * {@link Version#getVersion()}. POI 5.2.3 introduced the bug-51037 fix in
     * {@code XSSFRow.onDocumentWrite} that applies a default cell style to every
     * unstyled cell, causing {@code s="0"} to be emitted unconditionally —
     * diverging from POI 4.x ~ 5.2.2 output. Tests that assert DOM equality
     * against POI's output use this predicate to skip on lower versions.
     */
    public static boolean isPoi523OrLater() {
        final int[] v = _parseVersion(Version.getVersion());
        return v[0] > 5
                || (v[0] == 5 && v[1] > 2)
                || (v[0] == 5 && v[1] == 2 && v[2] >= 3);
    }

    private static int[] _parseVersion(final String version) {
        final String[] parts = version.split("\\.");
        return new int[]{
                _parseLeadingInt(parts.length > 0 ? parts[0] : "0"),
                _parseLeadingInt(parts.length > 1 ? parts[1] : "0"),
                _parseLeadingInt(parts.length > 2 ? parts[2] : "0"),
        };
    }

    /** Parses leading digits as int; tolerates suffixes like "5.5.1-beta1". */
    private static int _parseLeadingInt(final String s) {
        int end = 0;
        while (end < s.length() && Character.isDigit(s.charAt(end))) end++;
        return end == 0 ? 0 : Integer.parseInt(s.substring(0, end));
    }
}
