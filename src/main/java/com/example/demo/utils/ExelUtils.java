package com.example.demo.utils;

import com.example.demo.controller.model.FileItem;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Slf4j
@SuppressWarnings("unused")
public class ExelUtils {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    private static final SimpleDateFormat SIMPLE_DATE_FORMAT = new SimpleDateFormat("dd/MM/yyyy");

    public static void generateFileHeader(String filePath,
                                          String title,
                                          List<String> fileDescriptions,
                                          List<String> headers) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            int rowNum = 0;

            CellStyle titleStyle = workbook.createCellStyle();
            Font titleFont = workbook.createFont();
            titleFont.setFontHeightInPoints((short) 20);
            titleFont.setBold(true);
            titleStyle.setFont(titleFont);
            XSSFColor LIGHT_RED = new XSSFColor(new byte[]{(byte) 255, (byte) 160, (byte) 160}, null);
//            titleStyle.setFillForegroundColor(LIGHT_RED);
//            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Row titleRow = sheet.createRow(rowNum++);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue(title);
            titleCell.setCellStyle(titleStyle);

            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headers.size() - 1));

//            CellStyle descriptionStyle = workbook.createCellStyle();
//            descriptionStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
//            descriptionStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (String description : fileDescriptions) {
                Row descriptionRow = sheet.createRow(rowNum);
                Cell descriptionCell = descriptionRow.createCell(0);
                descriptionCell.setCellValue(description);
//                descriptionCell.setCellStyle(descriptionStyle);
                sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, headers.size() - 1));
                rowNum++;
            }

            Row headerRow = sheet.createRow(rowNum);
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);

            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
                cell.setCellStyle(headerStyle);
            }

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
            log.info("Excel file written successfully at: {}", filePath);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
    }

    public static FileItem loadFileAsResource(String filePath) {
        try {
            Path file = Paths.get(filePath);
            ByteArrayResource resource = new ByteArrayResource(Files.readAllBytes(file));
            return FileItem.builder()
                    .fileName(file.getFileName().toString())
                    .resource(resource)
                    .build();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            return null;
        }
    }

    public static void writeWorkBookToFile(Workbook workbook, String filename) {
        try (FileOutputStream fileOut = new FileOutputStream(filename)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            log.error("Error writing to file: {}", e.getMessage());
        }
    }

    public static void cleanUp(Path path) {
        try {
            Files.delete(path);
        } catch (IOException e) {
            log.error("Failed to delete file: {}", path);
        }
    }

    public static <T> void mapObjectToRow(Row row, T t, int headerLastRowNum, int rowIndex) {
        int column = 0;
        int headerHeight = rowIndex - headerLastRowNum - 1;
        row.createCell(column++).setCellValue(headerHeight);

        DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");

        Field[] fields = t.getClass().getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            try {
                Object value = field.get(t);
                if (value != null) {
                    switch (value.getClass().getSimpleName()) {
                        case "LocalDateTime" ->
                                row.createCell(column++).setCellValue(((LocalDateTime) value).format(DATE_FORMATTER));
                        case "LocalDate" ->
                                row.createCell(column++).setCellValue(((LocalDate) value).format(DATE_FORMATTER));
                        case "OffsetDateTime" ->
                                row.createCell(column++).setCellValue(((OffsetDateTime) value).format(DATE_FORMATTER));
                        case "Date" ->
                                row.createCell(column++).setCellValue(SIMPLE_DATE_FORMAT.format((Date) value));
                        case "String" ->
                                row.createCell(column++).setCellValue((String) value);
                        case "Integer", "Double", "Float", "Long", "Short" ->
                                row.createCell(column++).setCellValue(((Number) value).doubleValue());
                        case "BigDecimal" ->
                                row.createCell(column++).setCellValue(((BigDecimal) value).doubleValue());
                        case "Boolean" ->
                                row.createCell(column++).setCellValue((Boolean) value);
                        default ->
                                row.createCell(column++).setCellValue(value.toString());
                    }
                } else {
                    row.createCell(column++).setCellValue("");
                }
            } catch (IllegalAccessException e) {
                log.error("Error accessing field: {}", field.getName(), e);
            }
        }
    }

    public static <T> List<String> extractHeadersFromObjects(Class<T> t) {
        List<String> headers = new ArrayList<>();
        headers.add("NO");
        Field[] fields = t.getDeclaredFields();
        for (Field field : fields) {
            headers.add(field.getName().toUpperCase());
        }
        return headers;
    }

}
