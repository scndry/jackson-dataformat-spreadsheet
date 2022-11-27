package io.github.scndry.jackson.dataformat.spreadsheet.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.lang.ref.PhantomReference;
import java.lang.ref.ReferenceQueue;

@Slf4j
final class TempFileReference extends PhantomReference<Object> {

    private final String _pathname;

    TempFileReference(final String pathname, final Object referent, final ReferenceQueue<Object> q) {
        super(referent, q);
        _pathname = pathname;
    }

    public void clean() {
        if (log.isDebugEnabled()) {
            log.debug("Deleting temp file [{}]", _pathname);
        }
        FileUtils.deleteQuietly(new File(_pathname));
    }
}
