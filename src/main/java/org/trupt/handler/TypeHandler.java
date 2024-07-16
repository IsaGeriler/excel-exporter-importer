package org.trupt.handler;

import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;

public interface TypeHandler {
    void handle(Field field, Object instance, Cell cell) throws IllegalAccessException;
}