package io.github.scndry.jackson.dataformat.spreadsheet.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marks an API element (type or method) as incubating — added recently and
 * subject to non-breaking refinement based on user feedback. Backward-compatible
 * additions are expected; minor breaking changes remain possible in subsequent
 * minor releases until the API stabilizes through wide adoption.
 *
 * @since 1.6.0
 */
@Target({ElementType.TYPE, ElementType.METHOD})
@Retention(RetentionPolicy.CLASS)
@Documented
public @interface Incubating {
}
