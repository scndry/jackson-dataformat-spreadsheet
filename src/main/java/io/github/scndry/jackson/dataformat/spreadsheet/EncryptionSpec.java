package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.poi.poifs.crypt.ChainingMode;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.Incubating;

/**
 * Encryption strength preset for password-protected OOXML output, used with
 * {@code SheetOutput.withPassword(password, spec)}. {@link #strong()} is the
 * default when a password is set without an explicit spec; {@link #custom()}
 * opens a builder for finer control.
 */
@Incubating
public final class EncryptionSpec {

    /** Underlying cipher family. */
    public enum Cipher {
        AES_128(CipherAlgorithm.aes128),
        AES_192(CipherAlgorithm.aes192),
        AES_256(CipherAlgorithm.aes256);

        final CipherAlgorithm poi;

        Cipher(final CipherAlgorithm poi) { this.poi = poi; }
    }

    /** Hash used for the password-derivation key and integrity check. */
    public enum Hash {
        SHA_1(HashAlgorithm.sha1),
        SHA_256(HashAlgorithm.sha256),
        SHA_384(HashAlgorithm.sha384),
        SHA_512(HashAlgorithm.sha512);

        final HashAlgorithm poi;

        Hash(final HashAlgorithm poi) { this.poi = poi; }
    }

    /** Excel-version compatibility envelope. */
    public enum Compatibility {
        /** Agile encryption — Excel 2013 and newer. */
        EXCEL_2013_PLUS(EncryptionMode.agile),
        /** Standard encryption — Excel 2007/2010. */
        EXCEL_2007_2010(EncryptionMode.standard);

        final EncryptionMode poi;

        Compatibility(final EncryptionMode poi) { this.poi = poi; }
    }

    private final Compatibility compatibility;
    private final Cipher cipher;
    private final Hash hash;
    private final int keyBits;

    private EncryptionSpec(final Compatibility compatibility, final Cipher cipher,
                           final Hash hash, final int keyBits) {
        this.compatibility = compatibility;
        this.cipher = cipher;
        this.hash = hash;
        this.keyBits = keyBits;
    }

    /** AES-256 + SHA-512 (Excel 2013+). Default. */
    public static EncryptionSpec strong() {
        return new EncryptionSpec(Compatibility.EXCEL_2013_PLUS, Cipher.AES_256, Hash.SHA_512, 256);
    }

    /** AES-128 + SHA-512 (Excel 2013+). LibreOffice rejects this combination — use {@link #strong()} or {@link #legacy()} for LibreOffice compatibility. */
    public static EncryptionSpec balanced() {
        return new EncryptionSpec(Compatibility.EXCEL_2013_PLUS, Cipher.AES_128, Hash.SHA_512, 128);
    }

    /** AES-128 + SHA-1 (Excel 2007/2010). Opens in older Excel installs that reject agile mode. */
    public static EncryptionSpec legacy() {
        return new EncryptionSpec(Compatibility.EXCEL_2007_2010, Cipher.AES_128, Hash.SHA_1, 128);
    }

    /** AES-128 + SHA-256 (Excel 2013+). Lighter hash than {@link #balanced()}. */
    public static EncryptionSpec fast() {
        return new EncryptionSpec(Compatibility.EXCEL_2013_PLUS, Cipher.AES_128, Hash.SHA_256, 128);
    }

    /** Builder for custom combinations. Invalid combinations raise at {@link Builder#build()}. */
    public static Builder custom() { return new Builder(); }

    public Compatibility getCompatibility() { return compatibility; }
    public Cipher getCipher() { return cipher; }
    public Hash getHash() { return hash; }
    public int getKeyBits() { return keyBits; }

    // AES uses a fixed 16-byte block at every key size; POI expects this in EncryptionInfo#blockSize.
    private static final int AES_BLOCK_SIZE_BYTES = 16;

    /** Internal — used by the write path to build the POI {@link EncryptionInfo}. */
    public EncryptionInfo toEncryptionInfo() {
        // POI: standard encryption supports only ECB; agile uses CBC.
        final ChainingMode chaining = compatibility == Compatibility.EXCEL_2007_2010
                ? ChainingMode.ecb : ChainingMode.cbc;
        return new EncryptionInfo(compatibility.poi, cipher.poi, hash.poi,
                keyBits, AES_BLOCK_SIZE_BYTES, chaining);
    }

    @Override
    public String toString() {
        return "EncryptionSpec(" + compatibility + ", " + cipher + ", " + hash
                + ", " + keyBits + " bits)";
    }

    public static final class Builder {
        private Compatibility compatibility = Compatibility.EXCEL_2013_PLUS;
        private Cipher cipher = Cipher.AES_256;
        private Hash hash = Hash.SHA_512;
        private int keyBits = 256;
        private boolean keyBitsSet;

        private Builder() {}

        public Builder compatibility(final Compatibility c) {
            this.compatibility = c;
            return this;
        }

        public Builder cipher(final Cipher c) {
            this.cipher = c;
            if (!keyBitsSet) {
                this.keyBits = (c == Cipher.AES_256 ? 256 : c == Cipher.AES_192 ? 192 : 128);
            }
            return this;
        }

        public Builder hash(final Hash h) {
            this.hash = h;
            return this;
        }

        public Builder keyBits(final int bits) {
            if (bits != 128 && bits != 192 && bits != 256) {
                throw new IllegalArgumentException(
                        "keyBits must be 128, 192, or 256 (was " + bits + ")");
            }
            this.keyBits = bits;
            this.keyBitsSet = true;
            return this;
        }

        public EncryptionSpec build() {
            if (compatibility == Compatibility.EXCEL_2007_2010 && cipher == Cipher.AES_256) {
                throw new IllegalArgumentException(
                        "EXCEL_2007_2010 (standard mode) does not support AES-256");
            }
            return new EncryptionSpec(compatibility, cipher, hash, keyBits);
        }
    }
}
