package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.poi.openxml4j.opc.OPCPackage;

/**
 * Probes the runtime POI version for feature gates in tests.
 * The library targets POI 4.1.1+; some tests assert behavior that diverges
 * between POI 4.x and 5.x and are guarded with {@link #isPoi510OrLater()}.
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
}
