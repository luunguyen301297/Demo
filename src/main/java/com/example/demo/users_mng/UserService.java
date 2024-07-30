package com.example.demo.users_mng;

import com.example.demo.users_mng.utils.SimplePage;
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

    public SimplePage<User> getAll(Pageable pageable) {
        var users = userRepository.findAll();
        return new SimplePage<>(users, users.size(), pageable);
    }

    public User getById(@RequestParam Long id) {
        return userRepository.findById(id).orElse(null);
    }

}
