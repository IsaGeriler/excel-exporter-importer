package org.trupt.handler;

import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;

public class IntTypeHandler implements TypeHandler {
    @Override
    public void handle(Field field, Object instance, Cell cell) throws IllegalAccessException {
        switch (cell.getCellType()) {
            case NUMERIC -> field.set(instance, (int) cell.getNumericCellValue());
            case STRING -> field.set(instance, Integer.parseInt(cell.getStringCellValue()));
            default -> throw new IllegalArgumentException("Unsupported cell type");
        }
    }
}