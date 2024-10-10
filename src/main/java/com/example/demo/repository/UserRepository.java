package com.example.demo.repository;

import com.example.demo.model.UserEntity;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

@Repository
@SuppressWarnings("all")
public interface UserRepository extends JpaRepository<UserEntity, Long> {

    @Query(value = "select * from users ", nativeQuery = true)
    Page<UserEntity> findAll(Pageable pageable);

}
