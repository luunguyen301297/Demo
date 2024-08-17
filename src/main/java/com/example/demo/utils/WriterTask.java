package com.example.demo.utils;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.BlockingQueue;

@Slf4j
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public final class WriterTask<T> implements Runnable {

    BlockingQueue<List<T>> queue;
    String filePath;
    TriConsumer<Row, T, Integer> rowConsumer;

    @Override
    public void run() {
        try (
                FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(fis)
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            int newRowNum = lastRowNum + 1;

            List<T> data;
            while (!(data = queue.take()).isEmpty()) {
                for (T item : data) {
                    Row row = sheet.createRow(newRowNum++);
                    rowConsumer.accept(row, item, newRowNum);
                }
            }
            FileUtils.writeWorkBookToFile(workbook, filePath);
        } catch (IOException | InterruptedException e) {
            Thread.currentThread().interrupt();
            log.error("Writer thread interrupted: {}", e.getMessage());
        }
    }

}
