package org.trupt.utils;

import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.trupt.annotation.ExcelCellHeader;
import org.trupt.config.Log4j2Config;
import org.trupt.handler.TypeHandler;
import org.trupt.handler.TypeHandlerStorage;

import java.io.File;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class ImporterUtil {
    private static final Logger logger = Log4j2Config.getLogger(ExporterUtil.class);
    private final TypeHandlerStorage typeHandlerStorage = new TypeHandlerStorage();

    public <Type> List<Type> importFile(File file, Class<Type> type) {
        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            return StreamSupport.stream(sheet.spliterator(), false).skip(1) // Skip header
                    .takeWhile(row -> !isRowEmpty(row)).map(row -> {
                        try {
                            Type instance = type.getDeclaredConstructor().newInstance();
                            populateUserFromRow(instance, row, sheet);
                            return instance;
                        } catch (Exception e) {
                            logger.error("[INFO] An error occurred while setting field value: ", e);
                            return null;
                        }
                    }).filter(Objects::nonNull).collect(Collectors.toList());
        } catch (Exception e) {
            logger.error("[INFO] An error occurred while setting field value: ", e);
            return null;
        }
    }

    private boolean isRowEmpty(Row row) {
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

    private int getColumnIndexByName(Sheet sheet, String colName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(colName)) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalArgumentException("Column " + colName + " not found");
    }

    private void setFieldValue(Field field, Object instance, Cell cell) throws IllegalAccessException {
        TypeHandler handler = typeHandlerStorage.getHandler(field.getType());
        if (handler != null) {
            handler.handle(field, instance, cell);
        } else {
            throw new IllegalArgumentException("Unsupported field type: " + field.getType());
        }
    }

    private <Type> void populateUserFromRow(Type instance, Row row, Sheet sheet) {
        Field[] fields = instance.getClass().getDeclaredFields();
        for (Field field : fields) {
            try {
                if (field.isAnnotationPresent(ExcelCellHeader.class)) {
                    ExcelCellHeader header = field.getAnnotation(ExcelCellHeader.class);
                    int columnIndex = getColumnIndexByName(sheet, header.headerName().toUpperCase());
                    Cell cell = row.getCell(columnIndex);
                    if (cell == null && header.isRequired()) {
                        throw new RuntimeException(header.headerName()
                                + " column is required, cannot have NULL/BLANK values!");
                    } else if (cell == null) {
                        continue;
                    }
                    field.setAccessible(true);
                    setFieldValue(field, instance, cell);
                }
            } catch (Exception e) {
                logger.error("[INFO] An error occurred while setting field value: ", e);
            } finally {
                field.setAccessible(false);
            }
        }
    }
}