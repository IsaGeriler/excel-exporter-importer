package org.trupt.utils;

import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.trupt.annotation.ExcelCellHeader;
import org.trupt.config.Log4j2Config;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

public class ExporterUtil {
    private static final Logger logger = Log4j2Config.getLogger(ExporterUtil.class);

    public ByteArrayInputStream exportFile(List<?> list) {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = workbook.createSheet("Sheet1");
            XSSFRow headerRow = sheet.createRow(0);

            Class<?> clazz = list.get(0).getClass();
            Field[] fields = clazz.getDeclaredFields();

            int headerCellNo = 0;
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelCellHeader.class)) {
                    ExcelCellHeader column = field.getAnnotation(ExcelCellHeader.class);
                    XSSFCell cell = headerRow.createCell(headerCellNo);
                    cell.setCellValue(column.headerName());
                    makeRowBoldAndYellow(workbook, headerRow);

                    // Set initial column width based on header cell value
                    sheet.setColumnWidth(headerCellNo, (column.headerName().length() + 2) * 256);
                    headerCellNo++;
                }
            }

            int dataCellNo = 1;
            for (Object obj : list) {
                XSSFRow dataRow = sheet.createRow(dataCellNo++);
                int cellIndex = 0;
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ExcelCellHeader.class)) {
                        try {
                            field.setAccessible(true); // This allows us to bypass private and protected modifiers
                            Object value = field.get(obj);
                            XSSFCell cell = dataRow.createCell(cellIndex);

                            if (value != null) {
                                String cellValue;
                                if (value instanceof Number) {
                                    cell.setCellValue(((Number) value).doubleValue());
                                    cellValue = String.valueOf(((Number) value).doubleValue());
                                } else {
                                    cell.setCellValue(value.toString());
                                    cellValue = value.toString();
                                }

                                // Adjust column width based on cell value
                                int columnWidth = cellValue.length();
                                int currentWidth = sheet.getColumnWidth(cellIndex) / 256;
                                if (columnWidth + 2 > currentWidth) {
                                    sheet.setColumnWidth(cellIndex, (columnWidth + 2) * 256);
                                }
                            }
                        } catch (Exception e) {
                            field.setAccessible(false);
                            System.out.println("An error occurred while setting field value: "
                                    + Arrays.toString(e.getStackTrace()));
                        } finally {
                            field.setAccessible(false);
                        }
                        cellIndex++;
                    }
                }
            }

            // Create rows for SUM and AVERAGE calculations
            int sumRowNum = sheet.getLastRowNum() + 1;
            XSSFRow sumRow = sheet.createRow(sumRowNum);

            int avgRowNum = sheet.getLastRowNum() + 1;
            XSSFRow avgRow = sheet.createRow(avgRowNum);

            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelCellHeader.class)) {
                    ExcelCellHeader column = field.getAnnotation(ExcelCellHeader.class);
                    if (column.calculateSum()) {
                        String sumHeader = column.headerName();
                        calculateStatistic(sheet, sumHeader, "TOPLAM", "SUM", workbook, sumRow);
                    }
                    if (column.calculateAverage()) {
                        String avgHeader = column.headerName();
                        calculateStatistic(sheet, avgHeader, "ORTALAMA", "AVERAGEA", workbook, avgRow);
                    }
                }
            }
            workbook.write(byteArrayOutputStream);
            return new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
        } catch (Exception e) {
            logger.error("[INFO] An error occurred: ", e);
            return null;
        }
    }

    private void makeRowBoldAndYellow(Workbook workbook, Row row) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);

        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW1.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 0; i < row.getLastCellNum(); i++) {
            row.getCell(i).setCellStyle(cellStyle);
        }
    }

    private void makeCellBoldAndYellow(Workbook workbook, Cell cell) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);

        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW1.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
    }

    private int findTargetCellNum(XSSFRow headerRow, String cellHeader) {
        int targetCellNum = 0;
        try {
            for (int cellNum = 0; cellNum < headerRow.getLastCellNum(); cellNum++) {
                XSSFCell cell = headerRow.getCell(cellNum);
                if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().equals(cellHeader)) {
                    targetCellNum = cellNum;
                    break;
                }
            }
            return targetCellNum;
        } catch (Exception e) {
            logger.error("[INFO] An error occurred: ", e);
            return -1;
        }
    }

    private boolean isNonEmptyNumericCell(XSSFRow row, int targetCellNum) {
        if (row != null) {
            XSSFCell cell = row.getCell(targetCellNum);
            return ((cell != null) && (cell.getCellType() != CellType.BLANK) && (cell.getCellType() == CellType.NUMERIC));
        }
        return false;
    }

    private int firstNonEmptyCellNo(XSSFSheet sheet, int targetCellNum) {
        int firstDataRow = -1;
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            if (isNonEmptyNumericCell(row, targetCellNum)) {
                firstDataRow = rowNum + 1;
                break;
            }
        }
        return firstDataRow;
    }

    private int lastNonEmptyCellNo(XSSFSheet sheet, int targetCellNum) {
        int lastDataRow = -1;
        for (int rowNum = sheet.getLastRowNum(); rowNum >= 1; rowNum--) {
            XSSFRow row = sheet.getRow(rowNum);
            if (isNonEmptyNumericCell(row, targetCellNum)) {
                lastDataRow = rowNum + 1;
                break;
            }
        }
        return lastDataRow;
    }

    private void processCellHeaders(XSSFSheet sheet, XSSFRow headerRow, XSSFRow targetRow,
                                    String cellHeader, String formulaType) throws Exception {
        int targetCellNum = findTargetCellNum(headerRow, cellHeader);

        // Find the first and last non-empty cells in the column
        int firstDataRow = firstNonEmptyCellNo(sheet, targetCellNum);
        int lastDataRow = lastNonEmptyCellNo(sheet, targetCellNum);

        if (firstDataRow == -1 || lastDataRow == -1) {
            throw new Exception("No data found in column '" + cellHeader + "'.");
        }

        XSSFCell formulaCell = targetRow.createCell(targetCellNum);
        formulaCell.setCellFormula(formulaType + "(" + CellReference.convertNumToColString(targetCellNum)
                + firstDataRow + ":" + CellReference.convertNumToColString(targetCellNum) + lastDataRow + ")");
    }

    private void calculateStatistic(XSSFSheet sheet, String cellHeader, String label, String formula, Workbook workbook, XSSFRow targetRow) {
        try {
            // Add label only once in the first cell of the statistics row
            if (targetRow.getCell(0) == null) {
                targetRow.createCell(0).setCellValue(label);
                makeCellBoldAndYellow(workbook, targetRow.getCell(0));

                int columnWidth = targetRow.getCell(0).getStringCellValue().length();
                int currentWidth = sheet.getColumnWidth(0) / 256;
                if (columnWidth + 2 > currentWidth) {
                    sheet.setColumnWidth(0, (columnWidth + 2) * 256);
                }
            }

            XSSFRow headerRow = sheet.getRow(0); // Assuming the first row is the header row
            processCellHeaders(sheet, headerRow, targetRow, cellHeader, formula);
        } catch (Exception e) {
            System.out.println("An error occurred while setting field value: " + Arrays.toString(e.getStackTrace()));
        }
    }
}