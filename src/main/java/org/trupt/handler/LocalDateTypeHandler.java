package org.trupt.handler;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.Objects;

public class LocalDateTypeHandler implements TypeHandler {
    @Override
    public void handle(Field field, Object instance, Cell cell) throws IllegalAccessException {
        if (Objects.requireNonNull(cell.getCellType()) == CellType.STRING) {
            field.set(instance, LocalDate.parse(cell.getStringCellValue()));
        } else {
            throw new IllegalArgumentException("Unsupported cell type");
        }
    }
}