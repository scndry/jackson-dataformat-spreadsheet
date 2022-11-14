package support;

import java.io.File;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;

public interface FixtureAs {

    default InputStream fixtureAsStream(final String name) {
        final Path path = Paths.get("support/fixture").resolve(name);
        return getClass().getClassLoader().getResourceAsStream(path.toString());
    }

    default File fixtureAsFile(final String name) {
        final Path path = Paths.get("support/fixture").resolve(name);
        final URL url = getClass().getClassLoader().getResource(path.toString());
        if (url == null) return null;
        try {
            return new File(url.toURI());
        } catch (URISyntaxException e) {
            throw new IllegalStateException(e);
        }
    }
}
