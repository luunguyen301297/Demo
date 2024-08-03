package com.example.demo.pageable;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.experimental.FieldDefaults;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;

import java.util.List;
import java.util.stream.Collectors;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class SimplePage<T> {

    List<T> content;
    long totalElements;
    int page;
    int size;
    int totalPages;
    @Getter
    List<Sort.Order> sort;

    @JsonCreator
    public SimplePage(@JsonProperty("content") final List<T> content,
                      @JsonProperty("totalElements") final long totalElements,
                      @JsonProperty("page") final int page,
                      @JsonProperty("size") final int size,
                      @JsonProperty("sort") final List<String> sort) {
        this.content = content;
        this.totalElements = totalElements;
        this.page = page;
        this.size = size;
        this.sort = sort.stream()
                .map(s -> {
                    String[] parts = s.split(",");
                    return new Sort.Order(Sort.Direction.fromString(parts[1]), parts[0]);
                })
                .collect(Collectors.toList());
        this.totalPages = calculateTotalPages(totalElements, size);
    }

    public SimplePage(final List<T> content, final long totalElements, final Pageable pageable) {
        this.content = content;
        this.totalElements = totalElements;
        this.page = pageable.getPageNumber();
        this.size = pageable.getPageSize();
        this.sort = pageable.getSort().toList();
        this.totalPages = calculateTotalPages(totalElements, pageable.getPageSize());
    }

    @JsonProperty("content")
    public List<T> getContent() {
        return content;
    }

    @JsonProperty("page")
    public int getPage() {
        return page;
    }

    @JsonProperty("size")
    public int getSize() {
        return size;
    }

    @JsonProperty("totalElements")
    public long getTotalElements() {
        return totalElements;
    }

    @JsonProperty("totalPages")
    public int getTotalPages() {
        return totalPages;
    }

    @JsonProperty("number")
    public int getNumber() {
        return page;
    }

    @JsonProperty("sort")
    public List<String> getSortList() {
        return sort.stream()
                .map(order -> order.getProperty() + "," + order.getDirection().name())
                .collect(Collectors.toList());
    }

    private int calculateTotalPages(long totalElements, int size) {
        return (int) Math.ceil((double) totalElements / size);
    }

}
