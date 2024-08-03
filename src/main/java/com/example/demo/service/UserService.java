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
import org.springframework.web.bind.annotation.RequestParam;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ExecutorService;

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

    public UserEntity getById(@RequestParam Long id) {
        return userRepository.findById(id).orElse(null);
    }

    public FileItem exportToFile() {
        long startTime = System.currentTimeMillis();
        String filePath = String.format("User_Report_%s.xlsx",
                new SimpleDateFormat("ddMMyyyy").format(new Date()));
        final String title = "USER_REPORT";
        List<String> headers = Arrays.asList("NO", "ID", "USERNAME", "EMAIL", "CREATED_AT");
        List<String> inputRequests = new ArrayList<>();
        inputRequests.add("From: dd/MM/yyyy" + " - To: dd/MM/yyyy");

        FileUtils.generateFileHeader(filePath, title, headers, inputRequests);

        int page = 0;
        int pageSize = 100;
        boolean nextScan = true;
        while (nextScan) {
            Pageable pageable = PageRequest.of(page, pageSize);
            SimplePage<UserEntity> data = this.getAll(pageable);

            if (data == null || CollectionUtils.isEmpty(data.getContent()) || data.getContent().size() < pageSize) {
                nextScan = false;
            }
            assert data != null;
            this.appendData(filePath, data.getContent(), page, pageSize);
            page += 1;
        }

        log.info("Exec time : {}", System.currentTimeMillis() - startTime);
        return FileUtils.loadFileAsResource(filePath);
    }

    public void appendData(String filePath, List<UserEntity> users, Integer page, Integer pageSize) {
        try (
                FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(fis)
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            int newRowNum = lastRowNum + 1;

            int noOfRow = 0;
            for (UserEntity u : users) {
                noOfRow += 1;
                Row row = sheet.createRow(newRowNum++);
                int column = 0;

                row.createCell(column++).setCellValue(page * pageSize + noOfRow);
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
