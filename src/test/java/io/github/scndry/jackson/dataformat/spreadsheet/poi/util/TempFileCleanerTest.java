package io.github.scndry.jackson.dataformat.spreadsheet.poi.util;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;

class TempFileCleanerTest {

    ExecutorService executor;
    TempFileCleaner cleaner;
    Path path;

    @BeforeEach
    void setUp() throws Exception {
        executor = Executors.newSingleThreadExecutor();
        cleaner = new DefaultTempFileCleaner();
        path = Files.createTempFile("test", ".tmp");
        assertThat(path).exists();
    }

    @Test
    @SuppressWarnings("StatementWithEmptyBody")
    void test() throws Exception {
        File file = path.toFile(); // reference to file
        cleaner.register(file);

        final Callable<Object> waitIfExists = () -> {
            while (Files.exists(path)) { /* ignore */ }
            return null;
        };

        assertThat(executor.submit(waitIfExists)).failsWithin(1, TimeUnit.SECONDS);
        assertThat(path).exists();

        file = null; // remove reference
        System.gc(); // force gc
        assertThat(executor.submit(waitIfExists)).succeedsWithin(1, TimeUnit.SECONDS);
        assertThat(path).doesNotExist();
    }
}
