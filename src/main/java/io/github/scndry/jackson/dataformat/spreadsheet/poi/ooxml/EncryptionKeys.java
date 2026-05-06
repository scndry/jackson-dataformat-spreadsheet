package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.security.SecureRandom;
import java.util.Arrays;

import org.apache.commons.codec.binary.Hex;

/**
 * Random key generation for the H2 MVStore-backed shared strings cache.
 * H2's {@code MVStore.Builder.encryptionKey(char[])} expects hex-encoded
 * characters; this helper produces 32 lower-case hex characters from
 * 16 cryptographically random bytes.
 */
final class EncryptionKeys {

    private EncryptionKeys() {
    }

    static char[] generate() {
        final byte[] bytes = new byte[16];
        new SecureRandom().nextBytes(bytes);
        final char[] key = Hex.encodeHex(bytes);
        Arrays.fill(bytes, (byte) 0);
        return key;
    }
}
