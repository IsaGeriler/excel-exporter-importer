package org.trupt.handler;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Objects;

public class BigDecimalTypeHandler implements TypeHandler {
    @Override
    public void handle(Field field, Object instance, Cell cell) throws IllegalAccessException {
        if (Objects.requireNonNull(cell.getCellType()) == CellType.NUMERIC) {
            field.set(instance, BigDecimal.valueOf(cell.getNumericCellValue()));
        } else {
            throw new IllegalArgumentException("Unsupported cell type");
        }
    }
}