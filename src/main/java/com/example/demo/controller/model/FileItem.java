package com.example.demo.controller.model;

import lombok.*;
import org.springframework.core.io.Resource;

@Builder
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class FileItem {

    private Resource resource;
    private String fileName;

}
