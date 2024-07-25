package org.trupt.utils;

import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.trupt.annotation.ExcelCellHeader;
import org.trupt.config.Log4j2Config;
import org.trupt.handler.TypeHandler;
import org.trupt.handler.TypeHandlerStorage;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class ImporterUtil {
    private static final Logger logger = Log4j2Config.getLogger(ImporterUtil.class);
    private final TypeHandlerStorage typeHandlerStorage = new TypeHandlerStorage();

    public <Type> List<Type> importFile(File file, Class<Type> type) {
        if (file == null) {
            logger.error("[ERROR] File is null.");
            throw new IllegalArgumentException("File cannot be null.");
        }

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            logger.info("[INFO] Starting to import file: {}", file.getAbsolutePath());

            List<Type> result = StreamSupport.stream(sheet.spliterator(), false)
                    .skip(1) // Skip header
                    .takeWhile(row -> !isRowEmpty(row))
                    .map(row -> {
                        try {
                            Type instance = type.getDeclaredConstructor().newInstance();
                            populateUserFromRow(instance, row, sheet);
                            return instance;
                        } catch (InstantiationException e) {
                            logger.error("[ERROR] Error creating an instance of {}: ", type.getSimpleName(), e);
                            throw new RuntimeException("Error creating an instance of " + type.getSimpleName(), e);
                        } catch (IllegalAccessException e) {
                            logger.error("[ERROR] Illegal access while creating an instance of {}: ", type.getSimpleName(), e);
                            throw new RuntimeException("Illegal access while creating an instance of " + type.getSimpleName(), e);
                        } catch (NoSuchMethodException e) {
                            logger.error("[ERROR] No suitable constructor for {}: ", type.getSimpleName(), e);
                            throw new RuntimeException("No suitable constructor for " + type.getSimpleName(), e);
                        } catch (InvocationTargetException e) {
                            logger.error("[ERROR] Error invoking constructor for {}: ", type.getSimpleName(), e);
                            throw new RuntimeException("Error invoking constructor for " + type.getSimpleName(), e);
                        } catch (Exception e) {
                            logger.error("[ERROR] An unexpected error occurred while processing row: ", e);
                            throw new RuntimeException("Unexpected error occurred while processing row", e);
                        }
                    })
                    .filter(Objects::nonNull)
                    .collect(Collectors.toList());
            logger.info("[INFO] Finished importing file: {}, imported {} rows", file.getAbsolutePath(), result.size());
            return result;
        } catch (IOException e) {
            logger.error("[ERROR] An error occurred while reading the file: ", e);
            throw new RuntimeException("Error reading the file", e);
        } catch (Exception e) {
            logger.error("[ERROR] An error occurred while processing the file: ", e);
            throw new RuntimeException("Unexpected error occurred while processing the file", e);
        }
    }

    private boolean isRowEmpty(Row row) {
        for (Cell cell : row)
            if (cell.getCellType() != CellType.BLANK) return false;
        return true;
    }

    private int getColumnIndexByName(Sheet sheet, String colName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow)
            if (cell.getStringCellValue().equalsIgnoreCase(colName)) return cell.getColumnIndex();
        throw new IllegalArgumentException("Column " + colName + " not found");
    }

    private void setFieldValue(Field field, Object instance, Cell cell) throws IllegalAccessException {
        TypeHandler handler = typeHandlerStorage.getHandler(field.getType());
        if (handler == null) throw new IllegalArgumentException("Unsupported field type: " + field.getType());
        handler.handle(field, instance, cell);
    }

    private <Type> void populateUserFromRow(Type instance, Row row, Sheet sheet) {
        Field[] fields = instance.getClass().getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelCellHeader.class)) {
                ExcelCellHeader header = field.getAnnotation(ExcelCellHeader.class);
                int columnIndex = getColumnIndexByName(sheet, header.headerName());

                Cell cell = row.getCell(columnIndex);
                try {
                    if (cell == null && header.isRequired()) {
                        throw new IllegalArgumentException(header.headerName() + " column is required, cannot have NULL/BLANK values!");
                    } else if (cell != null) {
                        field.setAccessible(true);
                        setFieldValue(field, instance, cell);
                    }
                } catch (IllegalArgumentException e) {
                    logger.error("[ERROR] Required field missing: {}", header.headerName(), e);
                    throw e;
                } catch (IllegalAccessException e) {
                    logger.error("[ERROR] Illegal access while setting field value for field {}: ", field.getName(), e);
                    throw new RuntimeException("Error accessing field " + field.getName(), e);
                } catch (Exception e) {
                    logger.error("[ERROR] An unexpected error occurred while setting field value for field {}: ", field.getName(), e);
                    throw new RuntimeException("Unexpected error occurred while setting field value", e);
                } finally {
                    field.setAccessible(false);
                }
            }
        }
    }
}