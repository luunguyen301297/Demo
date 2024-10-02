package com.example.demo.utils;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.util.CollectionUtils;

import java.util.Collections;
import java.util.List;
import java.util.concurrent.BlockingQueue;

@Slf4j
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public final class DbReaderTask<T, ID> implements Runnable {

    JpaRepository<T, ID> repository;
    BlockingQueue<List<T>> queue;
    int queryBatchSize;

    @Override
    public void run() {
        int page = 0;
        boolean nextScan = true;
        Pageable pageable;
        List<T> data;

        try {
            while (nextScan) {
                pageable = PageRequest.of(page++, queryBatchSize);
                data = repository.findAll(pageable).getContent();

                if (CollectionUtils.isEmpty(data)) {
                    nextScan = false;
                } else {
                    queue.put(data);
                    if (data.isEmpty()) {
                        nextScan = false;
                    }
                }
            }
            queue.put(Collections.emptyList());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            log.error("Reader thread interrupted: {}", e.getMessage());
        }
    }

}
