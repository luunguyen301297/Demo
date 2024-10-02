package com.example.demo.utils;

import com.example.demo.controller.model.FileItem;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

@Slf4j
public class ExelUtils {

    public static void generateFileHeader(String filePath,
                                          String title,
                                          List<String> fileDescriptions,
                                          List<String> headers) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            int rowNum = 0;

            Row titleRow = sheet.createRow(rowNum++);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue(title);

            // Create a cell style with larger font for the title
            CellStyle titleStyle = workbook.createCellStyle();
            Font titleFont = workbook.createFont();
            titleFont.setFontHeightInPoints((short) 16);
            titleFont.setBold(true);
            titleStyle.setFont(titleFont);
            titleCell.setCellStyle(titleStyle);

            // Merge title cell across columns
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headers.size() - 1));

            for (String description : fileDescriptions) {
                Row descriptionRow = sheet.createRow(rowNum);
                descriptionRow.createCell(0).setCellValue(description);
                sheet.addMergedRegion(new CellRangeAddress(1, fileDescriptions.size() , 0, headers.size() - 1));
                rowNum++;
            }

            Row headerRow = sheet.createRow(rowNum);
            for (int i = 0; i < headers.size(); i++) {
                headerRow.createCell(i).setCellValue(headers.get(i));
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

}
