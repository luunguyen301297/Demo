package com.example.demo.controller;

import com.example.demo.pageable.SimplePage;
import com.example.demo.repository.entity.UserEntity;
import com.example.demo.service.UserService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.data.domain.*;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;

import java.util.Collections;
import java.util.List;

import static org.hamcrest.Matchers.hasSize;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.hamcrest.Matchers.is;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

@WebMvcTest(UserController.class)
public class UserControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @MockBean
    private UserService userService;

    private UserEntity user;
    private SimplePage<UserEntity> userPage;

    @BeforeEach
    public void setUp() {
        user = new UserEntity();
        user.setId(1L);
        user.setUsername("John Doe");

        Pageable pageable = PageRequest.of(0, 5, Sort.by(Sort.Order.asc("id")));
        userPage = new SimplePage<>(List.of(user), 1, pageable);
    }

    @Test
    public void testGetUsers() throws Exception {
        Pageable pageable = PageRequest.of(0, 5, Sort.by(Sort.Order.asc("id")));
        when(userService.getAll(pageable)).thenReturn(userPage);

        mockMvc.perform(get("/users")
                        .param("page", "0")
                        .param("size", "5")
                        .param("sort", "id,asc")
                        .contentType(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.content", hasSize(1)))
                .andExpect(jsonPath("$.content[0].id", is(1)))
                .andExpect(jsonPath("$.content[0].username", is("John Doe")))
                .andExpect(jsonPath("$.totalElements", is(1)))
                .andExpect(jsonPath("$.totalPages", is(1)))
                .andExpect(jsonPath("$.page", is(0)))
                .andExpect(jsonPath("$.size", is(5)));
    }

    @Test
    public void testGetUsers_EmptyList() throws Exception {
        Pageable pageable = PageRequest.of(0, 5, Sort.by(Sort.Order.asc("id")));
        SimplePage<UserEntity> emptyPage = new SimplePage<>(Collections.emptyList(), 0, pageable);

        when(userService.getAll(pageable)).thenReturn(emptyPage);

        mockMvc.perform(get("/users")
                        .param("page", "0")
                        .param("size", "5")
                        .param("sort", "id,asc")
                        .contentType(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.content", hasSize(0)))
                .andExpect(jsonPath("$.totalElements", is(0)))
                .andExpect(jsonPath("$.totalPages", is(0)))
                .andExpect(jsonPath("$.page", is(0)))
                .andExpect(jsonPath("$.size", is(5)));
    }

    @Test
    public void testGetUser() throws Exception {
        when(userService.getById(1L)).thenReturn(user);

        mockMvc.perform(get("/users/1")
                        .accept(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.id", is(1)))
                .andExpect(jsonPath("$.username", is("John Doe")));
    }

    @Test
    public void testGetUser_NotFound() throws Exception {
        when(userService.getById(999L)).thenReturn(null);

        mockMvc.perform(get("/users/999")
                        .accept(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(content().string(""));
    }

}
