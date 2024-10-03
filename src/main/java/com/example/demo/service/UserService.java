package com.example.demo.service;

import com.example.demo.controller.model.FileItem;
import com.example.demo.pageable.SimplePage;
import com.example.demo.repository.UserRepository;
import com.example.demo.repository.entity.UserEntity;
import com.example.demo.utils.ExelUtils;
import com.example.demo.utils.DbReaderTask;
import com.example.demo.utils.DbWriterTask;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
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

    static int QUERY_BATCH_SIZE = 1000;

    public SimplePage<UserEntity> getAll(Pageable pageable) {
        var users = userRepository.findAll(pageable);
        return new SimplePage<>(users.getContent(), users.getTotalElements(), pageable);
    }

    public UserEntity getById(Long id) {
        return userRepository.findById(id).orElse(null);
    }

    public FileItem exportFromDbToFile() {
        String filePath = String.format("User_Report_%s.xlsx",
                new SimpleDateFormat("ddMMyyyy").format(new Date()));
        final String title = "USER_REPORT";
        List<String> headers = ExelUtils.extractHeadersFromObjects(UserEntity.class);
        List<String> fileDescriptions = new ArrayList<>();
        fileDescriptions.add("From: dd/MM/yyyy" + " - To: dd/MM/yyyy");

        ExelUtils.generateFileHeader(filePath, title, fileDescriptions, headers);

        BlockingQueue<List<UserEntity>> queue = new LinkedBlockingQueue<>(10);

        CompletableFuture<Void> readFuture = CompletableFuture
                .runAsync(new DbReaderTask<>(userRepository, queue, QUERY_BATCH_SIZE), executorService);

        CompletableFuture<Void> writeFuture = CompletableFuture
                .runAsync(new DbWriterTask<>(queue, filePath, ExelUtils::mapObjectToRow), executorService);

        CompletableFuture<Void> combinedFuture = CompletableFuture.allOf(readFuture, writeFuture);

        combinedFuture.thenRun(() -> log.info("Completed write to file: {}", filePath)).join();

        return ExelUtils.loadFileAsResource(filePath);
    }

}
