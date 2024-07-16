package org.trupt.handler;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

public class TypeHandlerStorage {
    private final Map<Class<?>, TypeHandler> handlers = new HashMap<>();

    public TypeHandlerStorage() {
        handlers.put(int.class, new IntTypeHandler());
        handlers.put(Integer.class, new IntTypeHandler());
        handlers.put(double.class, new DoubleTypeHandler());
        handlers.put(Double.class, new DoubleTypeHandler());
        handlers.put(String.class, new StringTypeHandler());
        handlers.put(BigDecimal.class, new BigDecimalTypeHandler());
        handlers.put(LocalDate.class, new LocalDateTypeHandler());
    }

    public TypeHandler getHandler(Class<?> type) {
        return handlers.get(type);
    }
}