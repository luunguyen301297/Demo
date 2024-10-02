package com.example.demo.controller;

import com.example.demo.controller.model.FileItem;
import com.example.demo.pageable.SimplePage;
import com.example.demo.repository.entity.UserEntity;
import com.example.demo.service.UserService;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import org.springframework.core.io.Resource;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.data.web.PageableDefault;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/users")
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class UserController {

    UserService userService;

    @GetMapping("")
    SimplePage<UserEntity> getUsers(
            @PageableDefault(size = 5, sort = "id", direction = Sort.Direction.ASC) Pageable pageable)
    {
        return userService.getAll(pageable);
    }

    @GetMapping("/{id}")
    UserEntity getUser(@PathVariable long id) {
        return userService.getById(id);
    }

    @GetMapping(value = "/export")
    ResponseEntity<Resource> reportDashboardFindExportToFile() {
        FileItem fileItem = userService.exportFromDbToFile();
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        "attachment; filename=\"" + fileItem.getFileName() + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(fileItem.getResource());
    }

}
