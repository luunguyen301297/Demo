package com.example.demo.users_mng.service;

import com.example.demo.users_mng.pageable.SimplePage;
import com.example.demo.users_mng.repository.UserRepository;
import com.example.demo.users_mng.repository.entity.UserEntity;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import org.springframework.data.domain.Pageable;
import org.springframework.stereotype.Service;
import org.springframework.web.bind.annotation.RequestParam;

@Service
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class UserService {

    UserRepository userRepository;

    public SimplePage<UserEntity> getAll(Pageable pageable) {
        var users = userRepository.findAll(pageable);
        return new SimplePage<>(users.getContent(), users.getTotalElements(), pageable);
    }

    public UserEntity getById(@RequestParam Long id) {
        return userRepository.findById(id).orElse(null);
    }

}
