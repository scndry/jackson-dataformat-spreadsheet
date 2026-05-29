package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.security.GeneralSecurityException;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.crypt.temp.EncryptedTempData;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import io.github.scndry.jackson.dataformat.spreadsheet.EncryptionSpec;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.POICompat;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;

/**
 * OOXML password protection helpers for the read and write paths.
 * Decrypts an encrypted source to an owner-only temp file before it
 * re-enters the regular read pipeline, and wraps a write target so the
 * plain OOXML output is encrypted into the final destination on close.
 */
@Slf4j
public final class OoxmlEncryption {

    private OoxmlEncryption() {
    }

    /**
     * Decrypts an encrypted OOXML source to an owner-only temp file and
     * returns a fresh {@link SheetInput} referencing it.
     */
    @SuppressWarnings("unchecked")
    public static SheetInput<?> decrypt(final SheetInput<?> src) throws IOException {
        final File encryptedFile;
        final boolean spooledFromStream;
        if (src.isFile()) {
            encryptedFile = ((SheetInput<File>) src).getRaw();
            spooledFromStream = false;
        } else {
            encryptedFile = _spoolToSecureTempFile(
                    ((SheetInput<InputStream>) src).getRaw(),
                    "jackson-spreadsheet-encrypted-");
            spooledFromStream = true;
        }

        File decrypted = null;
        try (POIFSFileSystem fs = _openEncryptedPOIFS(encryptedFile, src)) {
            final EncryptionInfo info;
            try {
                info = new EncryptionInfo(fs);
            } catch (IOException e) {
                throw new EncryptedDocumentException(
                        "Source is not an encrypted OOXML document (password was supplied"
                                + " but no EncryptionInfo entry found): " + src);
            } catch (IllegalArgumentException e) {
                // POI throws this for XLS — keep the type, rewrap with a guide-oriented message.
                throw new IllegalArgumentException(
                        "Password protection is supported for OOXML (XLSX) sources only: " + src, e);
            }
            final Decryptor d = Decryptor.getInstance(info);
            try {
                if (!d.verifyPassword(src.getPassword())) {
                    throw new EncryptedDocumentException(
                            "Invalid password for encrypted spreadsheet source: " + src);
                }
            } catch (GeneralSecurityException e) {
                throw new IOException("Failed to verify password for " + src, e);
            }
            decrypted = POICompat.createSecureTempFile(
                    "jackson-spreadsheet-decrypted-", ".xlsx").toFile();
            try (InputStream plain = d.getDataStream(fs);
                 FileOutputStream out = new FileOutputStream(decrypted)) {
                final byte[] buf = new byte[8192];
                int n;
                while ((n = plain.read(buf)) != -1) {
                    out.write(buf, 0, n);
                }
            } catch (GeneralSecurityException e) {
                throw new IOException("Failed to read decrypted data stream for " + src, e);
            }
        } catch (IOException | RuntimeException e) {
            if (decrypted != null) {
                try { POICompat.releaseTempFile(decrypted.toPath()); }
                catch (IOException cleanup) { e.addSuppressed(cleanup); }
            }
            throw e;
        } finally {
            if (spooledFromStream) {
                try { POICompat.releaseTempFile(encryptedFile.toPath()); }
                catch (IOException ignore) { /* best effort */ }
            }
        }
        if (log.isDebugEnabled()) {
            log.debug("Decrypted encrypted source to temp file: {}", decrypted);
        }
        return src.isNamed()
                ? SheetInput.source(decrypted, src.getName())
                : SheetInput.source(decrypted, src.getIndex());
    }

    /**
     * Wraps an encrypted target — plain OOXML is streamed into an
     * {@link EncryptedTempData} and encrypted on close.
     */
    public static SheetOutput<OutputStream> wrapTarget(final SheetOutput<?> out) throws IOException {
        final EncryptedTempData tempData = new EncryptedTempData();
        final EncryptionSpec spec = out.getEncryption() != null
                ? out.getEncryption() : EncryptionSpec.strong();
        final OutputStream tempStream;
        try {
            tempStream = new EncryptOnCloseOutputStream(tempData, out, out.getPassword(), spec);
        } catch (IOException | RuntimeException e) {
            tempData.dispose();
            throw e;
        }
        return out.isNamed()
                ? SheetOutput.target(tempStream, out.getName())
                : SheetOutput.target(tempStream);
    }

    @SuppressWarnings("unchecked")
    static void encryptToTarget(final EncryptedTempData tempData, final SheetOutput<?> out,
                                final String password, final EncryptionSpec spec) throws IOException {
        // File targets: write to a sibling temp, then atomic rename — partial bytes never reach the target.
        final File finalTarget = out.isFile() ? ((SheetOutput<File>) out).getRaw() : null;
        final File siblingTemp = finalTarget != null
                ? new File(finalTarget.getAbsoluteFile().getParentFile(),
                        finalTarget.getName() + ".enc-" + Long.toUnsignedString(System.nanoTime(), 36) + ".tmp")
                : null;
        final OutputStream target = (siblingTemp != null)
                ? Files.newOutputStream(siblingTemp.toPath())
                : ((SheetOutput<OutputStream>) out).getRaw();
        boolean writeFinalized = false;
        IOException primary = null;
        try {
            try (InputStream plainIn = tempData.getInputStream();
                 POIFSFileSystem fs = new POIFSFileSystem();
                 OPCPackage opc = OPCPackage.open(plainIn)) {
                final EncryptionInfo info = spec.toEncryptionInfo();
                final Encryptor enc = Encryptor.getInstance(info);
                try {
                    enc.confirmPassword(password);
                    // OPCPackage.save() only finishes the zip wrap, never closes;
                    // explicit close commits EncryptionInfo / EncryptedPackage to the POIFS root.
                    try (OutputStream encStream = enc.getDataStream(fs)) {
                        opc.save(encStream);
                    }
                } catch (GeneralSecurityException e) {
                    throw new IOException("Failed to encrypt spreadsheet output", e);
                }
                fs.writeFilesystem(target);
                writeFinalized = true;
            } catch (InvalidFormatException e) {
                throw new IOException("Failed to open temp data for encryption", e);
            }
        } catch (IOException e) {
            primary = e;
            throw e;
        } finally {
            if (siblingTemp != null) {
                try {
                    target.close();
                } catch (IOException closeEx) {
                    if (primary != null) primary.addSuppressed(closeEx);
                    else throw closeEx;
                }
                if (writeFinalized) {
                    try {
                        Files.move(siblingTemp.toPath(), finalTarget.toPath(),
                                StandardCopyOption.ATOMIC_MOVE,
                                StandardCopyOption.REPLACE_EXISTING);
                    } catch (AtomicMoveNotSupportedException ex) {
                        // Cross-filesystem — non-atomic, but partial state stays in sibling.
                        Files.move(siblingTemp.toPath(), finalTarget.toPath(),
                                StandardCopyOption.REPLACE_EXISTING);
                    }
                } else {
                    // Mid-write failure — delete sibling, leave target untouched.
                    if (siblingTemp.exists() && !siblingTemp.delete() && log.isWarnEnabled()) {
                        log.warn("Failed to delete partial encrypted sibling: {}", siblingTemp);
                    }
                }
            }
        }
    }

    private static POIFSFileSystem _openEncryptedPOIFS(final File encryptedFile, final SheetInput<?> src)
            throws IOException {
        try {
            return new POIFSFileSystem(encryptedFile, true);
        } catch (OfficeXmlFileException e) {
            // Plain OOXML (zip) reaches here as OfficeXmlFileException — POIFSFileSystem only opens OLE2.
            throw new EncryptedDocumentException(
                    "Source is not an encrypted OOXML document (password was supplied"
                            + " but the file is a plain OOXML package): " + src);
        }
    }

    private static File _spoolToSecureTempFile(final InputStream in, final String prefix) throws IOException {
        final File temp = POICompat.createSecureTempFile(prefix, ".tmp").toFile();
        try (FileOutputStream out = new FileOutputStream(temp)) {
            final byte[] buf = new byte[8192];
            int n;
            while ((n = in.read(buf)) != -1) {
                out.write(buf, 0, n);
            }
        } catch (IOException e) {
            try { POICompat.releaseTempFile(temp.toPath()); }
            catch (IOException cleanup) { e.addSuppressed(cleanup); }
            throw e;
        }
        return temp;
    }
}
