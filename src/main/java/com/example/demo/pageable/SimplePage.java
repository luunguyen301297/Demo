package com.example.demo.pageable;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AccessLevel;
import lombok.experimental.FieldDefaults;
import org.springframework.data.domain.Pageable;

import java.util.List;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class SimplePage<T> {

    List<T> content;
    long totalElements;
    int page;
    int size;
    int totalPages;

    @JsonCreator
    public SimplePage(@JsonProperty("content") final List<T> content,
                      @JsonProperty("totalElements") final long totalElements,
                      @JsonProperty("page") final int page,
                      @JsonProperty("size") final int size) {
        this.content = content;
        this.totalElements = totalElements;
        this.page = page;
        this.size = size;
        this.totalPages = calculateTotalPages(totalElements, size);
    }

    public SimplePage(final List<T> content, final long totalElements, final Pageable pageable) {
        this.content = content;
        this.totalElements = totalElements;
        this.page = pageable.getPageNumber();
        this.size = pageable.getPageSize();
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

    private int calculateTotalPages(long totalElements, int size) {
        return (int) Math.ceil((double) totalElements / size);
    }

}
