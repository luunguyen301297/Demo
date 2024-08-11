package com.example.demo.service;

import com.example.demo.controller.model.FileItem;
import com.example.demo.pageable.SimplePage;
import com.example.demo.repository.UserRepository;
import com.example.demo.repository.entity.UserEntity;
import com.example.demo.utils.FileUtils;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;

@Service
@Slf4j
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class UserService {

    UserRepository userRepository;

    ExecutorService executorService;

    public SimplePage<UserEntity> getAll(Pageable pageable) {
        var users = userRepository.findAll(pageable);
        return new SimplePage<>(users.getContent(), users.getTotalElements(), pageable);
    }

    public UserEntity getById(Long id) {
        return userRepository.findById(id).orElse(null);
    }

    public FileItem exportToFile() {
        long start = System.currentTimeMillis();
        String filePath = String.format("User_Report_%s.xlsx",
                new SimpleDateFormat("ddMMyyyy").format(new Date()));
        final String title = "USER_REPORT";
        List<String> headers = Arrays.asList("NO", "ID", "USERNAME", "EMAIL", "CREATED_AT");
        List<String> inputRequests = new ArrayList<>();
        inputRequests.add("From: dd/MM/yyyy" + " - To: dd/MM/yyyy");

        FileUtils.generateFileHeader(filePath, title, headers, inputRequests);

        // Transfer data between the reader and writer threads
        BlockingQueue<List<UserEntity>> queue = new LinkedBlockingQueue<>(10);

        // Reader task
        int pageSize = 500;
        Runnable readTask = () -> {
            int page = 0;
            boolean nextScan = true;
            try {
                while (nextScan) {
                    Pageable pageable = PageRequest.of(page, pageSize);
                    SimplePage<UserEntity> data = this.getAll(pageable);

                    if (data == null || CollectionUtils.isEmpty(data.getContent()) || data.getContent().size() < pageSize) {
                        nextScan = false;
                    }
                    if (data != null && !CollectionUtils.isEmpty(data.getContent())) {
                        queue.put(data.getContent());
                    }
                    page += 1;
                }
                queue.put(Collections.emptyList());
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                log.error("Reader thread interrupted: {}", e.getMessage());
            }
        };

        // Writer task (pass page and pageSize to appendData)
        Runnable writeTask = () -> {
            int page = 0;
            try {
                List<UserEntity> users;
                while (!(users = queue.take()).isEmpty()) {
                    this.appendData(filePath, users, page, pageSize);
                    page++;
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                log.error("Writer thread interrupted: {}", e.getMessage());
            }
        };

        // Execute the tasks concurrently
        executorService.execute(readTask);
        executorService.execute(writeTask);

        // Shutdown the executor service after tasks completion
        executorService.shutdown();
        try {
            if (!executorService.awaitTermination(60, TimeUnit.MINUTES)) {
                executorService.shutdownNow();
            }
        } catch (InterruptedException e) {
            executorService.shutdownNow();
            Thread.currentThread().interrupt();
        }
        System.err.println(System.currentTimeMillis() - start);
        return FileUtils.loadFileAsResource(filePath);
    }

    private void appendData(String filePath, List<UserEntity> users, int page, int pageSize) {
        try (
                FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(fis)
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            int newRowNum = lastRowNum + 1;

            for (int i = 0; i < users.size(); i++) {
                UserEntity u = users.get(i);
                Row row = sheet.createRow(newRowNum++);
                int column = 0;
                int globalRowNumber = page * pageSize + i + 1;

                row.createCell(column++).setCellValue(globalRowNumber);
                row.createCell(column++).setCellValue(u.getId() != null ? String.valueOf(u.getId()) : "");
                row.createCell(column++).setCellValue(u.getUsername() != null ? u.getUsername() : "");
                row.createCell(column++).setCellValue(u.getEmail() != null ? u.getEmail() : "");
                row.createCell(column).setCellValue(u.getCreatedAt() != null ? u.getCreatedAt().toString() : "");
            }
            FileUtils.writeWorkBookToFile(workbook, filePath);
        } catch (IOException e) {
            log.error("Error reading workbook: {}", e.getMessage());
        }
    }

}
