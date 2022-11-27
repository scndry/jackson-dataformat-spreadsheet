package io.github.scndry.jackson.dataformat.spreadsheet.poi.util;

import lombok.extern.slf4j.Slf4j;

import java.lang.ref.ReferenceQueue;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.atomic.AtomicLong;

@Slf4j
public final class DefaultTempFileCleaner implements TempFileCleaner, Runnable {

    private static final AtomicLong ID = new AtomicLong();
    private static final ThreadFactory THREAD_FACTORY = r -> {
        final Thread t = new Thread(r);
        t.setName("temp-file-cleaner-" + ID.getAndIncrement());
        t.setDaemon(true);
        t.setPriority(Thread.MIN_PRIORITY);
        return t;
    };

    private final ReferenceQueue<Object> _queue = new ReferenceQueue<>();
    private final Set<TempFileReference> _refs = ConcurrentHashMap.newKeySet();
    private volatile boolean _started;
    private volatile boolean _shutdown;
    private Thread _thread;

    @Override
    public void register(final String pathname, final Object referent) {
        if (_shutdown) {
            throw new IllegalStateException(TempFileCleaner.class.getSimpleName() + " already shutdown");
        }
        if (log.isDebugEnabled()) {
            log.debug("Registering reference {} for [{}]", referent, pathname);
        }
        _refs.add(new TempFileReference(pathname, referent, _queue));
        if (!_started) {
            _started = true;
            _thread = THREAD_FACTORY.newThread(this);
            if (log.isDebugEnabled()) {
                log.debug("Starting thread [{}]", _thread.getName());
            }
            _thread.start();
        }
    }

    @Override
    public void run() {
        boolean interrupted = false;
        try {
            while (!_refs.isEmpty()) {
                try {
                    final TempFileReference ref = (TempFileReference) _queue.remove(60 * 1000L);
                    if (ref != null) {
                        _refs.remove(ref);
                        ref.clean();
                    }
                } catch (InterruptedException e) {
                    interrupted = true;
                }
            }
        } finally {
            if (interrupted) {
                Thread.currentThread().interrupt();
            }
            _thread = null;
            _started = false;
        }
    }

    @Override
    public void shutdown() {
        _shutdown = true;
        Thread t;
        if ((t = _thread) != null && !t.isInterrupted()) {
            t.interrupt();
        }
    }

    @Override
    public boolean isShutdown() {
        return _shutdown;
    }
}
