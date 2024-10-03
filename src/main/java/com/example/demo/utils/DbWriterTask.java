package com.example.demo.utils;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.BlockingQueue;

@Slf4j
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public final class DbWriterTask<T> implements Runnable {

    BlockingQueue<List<T>> queue;
    String filePath;
    QuadConsumer<Row, T, Integer, Integer> rowConsumer;

    @Override
    public void run() {
        try (
                FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(fis)
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            int headerLastRowNum = sheet.getLastRowNum();
            int newRowNum = headerLastRowNum + 1;

            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);

            List<T> data;
            while (!(data = queue.take()).isEmpty()) {
                for (T item : data) {
                    Row row = sheet.createRow(newRowNum++);
                    rowConsumer.accept(row, item, headerLastRowNum, newRowNum);
                    for (Cell cell : row) {
                        cell.setCellStyle(borderStyle);
                    }
                }
            }
            int maxCell = sheet.getRow(headerLastRowNum + 1).getLastCellNum();
            for (int i = 0; i < maxCell; i++) {
                sheet.autoSizeColumn(i);
            }
            ExelUtils.writeWorkBookToFile(workbook, filePath);
        } catch (IOException | InterruptedException e) {
            Thread.currentThread().interrupt();
            log.error("Writer thread interrupted: {}", e.getMessage());
        }
    }

}
