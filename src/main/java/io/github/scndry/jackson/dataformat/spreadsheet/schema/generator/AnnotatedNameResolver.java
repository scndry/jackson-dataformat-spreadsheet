package io.github.scndry.jackson.dataformat.spreadsheet.schema.generator;

import com.fasterxml.jackson.databind.BeanProperty;
import lombok.RequiredArgsConstructor;

import java.lang.annotation.Annotation;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;
import java.util.function.Function;

@RequiredArgsConstructor
public final class AnnotatedNameResolver<A extends Annotation> implements ColumnNameResolver {

    private final Class<A> _type;
    private final Function<A, String> _nameMapper;

    public static <A extends Annotation, T>
    AnnotatedNameResolver<A> forValue(final Class<A> type, final Function<T, String> nameMapper) {
        return forAttribute(type, "value", nameMapper);
    }

    public static <A extends Annotation, T>
    AnnotatedNameResolver<A> forAttribute(final Class<A> type, final String attribute,
                                          final Function<T, String> nameMapper) {
        return new AnnotatedNameResolver<>(type, annotation -> {
            try {
                final Method method = type.getMethod(attribute);
                return nameMapper.apply(invoke(method, annotation));
            } catch (Throwable e) {
                throw new IllegalStateException(e);
            }
        });
    }

    @SuppressWarnings("unchecked")
    private static <A extends Annotation, T> T invoke(final Method method, final A annotation) throws Throwable {
        if (Proxy.isProxyClass(annotation.getClass())) {
            final InvocationHandler handler = Proxy.getInvocationHandler(annotation);
            return (T) handler.invoke(annotation, method, new Object[0]);
        }
        return (T) method.invoke(annotation);
    }

    @Override
    public String resolve(final BeanProperty prop) {
        return _nameMapper.apply(prop.getAnnotation(_type));
    }
}
