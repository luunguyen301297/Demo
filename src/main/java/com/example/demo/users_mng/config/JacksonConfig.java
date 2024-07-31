package com.example.demo.users_mng.config;

import com.example.demo.users_mng.pageable.SimplePage;
import com.example.demo.users_mng.pageable.SimplePageSerializer;
import com.fasterxml.jackson.databind.module.SimpleModule;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class JacksonConfig {

    @Bean
    public SimpleModule simplePageModule() {
        SimpleModule module = new SimpleModule();
        module.addSerializer(SimplePage.class, new SimplePageSerializer());
        return module;
    }

}
