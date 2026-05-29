package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.FilterOutputStream;
import java.io.IOException;

import org.apache.poi.poifs.crypt.temp.EncryptedTempData;

import io.github.scndry.jackson.dataformat.spreadsheet.EncryptionSpec;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput;

/**
 * Streams plain OOXML bytes into an {@link EncryptedTempData} and, on
 * close, encrypts them into the final {@link SheetOutput} destination.
 */
final class EncryptOnCloseOutputStream extends FilterOutputStream {

    private final EncryptedTempData _tempData;
    private final SheetOutput<?> _destination;
    private final String _password;
    private final EncryptionSpec _spec;
    private boolean _closed;

    EncryptOnCloseOutputStream(final EncryptedTempData tempData,
                               final SheetOutput<?> destination,
                               final String password,
                               final EncryptionSpec spec) throws IOException {
        super(tempData.getOutputStream());
        _tempData = tempData;
        _destination = destination;
        _password = password;
        _spec = spec;
    }

    @Override
    public void close() throws IOException {
        if (_closed) return;
        _closed = true;
        super.close();
        try {
            OoxmlEncryption.encryptToTarget(_tempData, _destination, _password, _spec);
        } finally {
            _tempData.dispose();
        }
    }
}
