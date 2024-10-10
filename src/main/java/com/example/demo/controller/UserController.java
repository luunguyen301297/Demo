package com.example.demo.controller;

import com.example.demo.model.FileItem;
import com.example.demo.service.UserService;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/users")
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class UserController {

    UserService userService;

    @GetMapping(value = "/export")
    ResponseEntity<Resource> reportDashboardFindExportToFile() {
        FileItem fileItem = userService.exportFromDbToFile();
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        "attachment; filename=\"" + fileItem.getFileName() + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(fileItem.getResource());
    }

    @PostMapping(path = "/import", consumes = {"multipart/form-data"})
    ResponseEntity<Resource> importResources(@RequestParam("file") MultipartFile file) throws IOException {
        FileItem fileItem = userService.saveWithErrorReport(file);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        "attachment; filename=\"" + fileItem.getFileName() + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(fileItem.getResource());
    }

}
