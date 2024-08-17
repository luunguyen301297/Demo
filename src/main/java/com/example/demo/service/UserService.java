package com.example.demo.service;

import com.example.demo.controller.model.FileItem;
import com.example.demo.pageable.SimplePage;
import com.example.demo.repository.UserRepository;
import com.example.demo.repository.entity.UserEntity;
import com.example.demo.utils.FileUtils;
import com.example.demo.utils.ReaderTask;
import com.example.demo.utils.WriterTask;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.data.domain.Pageable;
import org.springframework.stereotype.Service;

import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.*;

@Service
@Slf4j
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class UserService {

    UserRepository userRepository;

    ExecutorService executorService;

    static int QUERY_BATCH_SIZE = 500;

    public SimplePage<UserEntity> getAll(Pageable pageable) {
        var users = userRepository.findAll(pageable);
        return new SimplePage<>(users.getContent(), users.getTotalElements(), pageable);
    }

    public UserEntity getById(Long id) {
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

        BlockingQueue<List<UserEntity>> queue = new LinkedBlockingQueue<>(10);

        CompletableFuture<Void> readFuture = CompletableFuture
                .runAsync(new ReaderTask<>(userRepository, queue, QUERY_BATCH_SIZE), executorService);

        CompletableFuture<Void> writeFuture = CompletableFuture
                .runAsync(new WriterTask<>(queue, filePath, this::mapUserEntityToRow), executorService);

        CompletableFuture<Void> combinedFuture = CompletableFuture.allOf(readFuture, writeFuture);

        combinedFuture.thenRun(() -> {
            log.info("Completed write to file: {}", filePath);
            log.info(String.valueOf(System.currentTimeMillis() - startTime));
        }).join();

        return FileUtils.loadFileAsResource(filePath);
    }

    private void mapUserEntityToRow(Row row, UserEntity user, int rowIndex) {
        int column = 0;
        row.createCell(column++).setCellValue(rowIndex - 3);
        row.createCell(column++).setCellValue(user.getId() != null ? String.valueOf(user.getId()) : "");
        row.createCell(column++).setCellValue(user.getUsername() != null ? user.getUsername() : "");
        row.createCell(column++).setCellValue(user.getEmail() != null ? user.getEmail() : "");
        row.createCell(column).setCellValue(user.getCreatedAt() != null ? user.getCreatedAt().toString() : "");
    }

}
