package com.example.demo.service;

import com.example.demo.pageable.SimplePage;
import com.example.demo.repository.UserRepository;
import com.example.demo.repository.entity.UserEntity;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageImpl;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;

import java.util.List;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.mockito.Mockito.when;

@ExtendWith(MockitoExtension.class)
public class UserServiceTest {

    @Mock
    private UserRepository userRepository;

    @InjectMocks
    private UserService userService;

    private UserEntity user;
    private Page<UserEntity> userPage;

    @BeforeEach
    public void setUp() {
        user = new UserEntity();
        user.setId(1L);
        user.setUsername("John Doe");

        userPage = new PageImpl<>(List.of(user));
    }

    @Test
    public void testGetAll() {
        Pageable pageable = PageRequest.of(0, 10);
        when(userRepository.findAll(pageable)).thenReturn(userPage);

        SimplePage<UserEntity> result = userService.getAll(pageable);

        assertEquals(1, result.getContent().size());
        assertEquals(1L, result.getContent().get(0).getId());
        assertEquals("John Doe", result.getContent().get(0).getUsername());
        assertEquals(1, result.getTotalElements());
    }

    @Test
    public void testGetById() {
        when(userRepository.findById(1L)).thenReturn(Optional.of(user));

        UserEntity result = userService.getById(1L);

        assertNotNull(result);
        assertEquals(1L, result.getId());
        assertEquals("John Doe", result.getUsername());
    }

    @Test
    public void testGetById_NotFound() {
        when(userRepository.findById(1L)).thenReturn(Optional.empty());
        UserEntity result = userService.getById(1L);
        assertNull(result);
    }

}
