package com.vivo.sprinbootpoi;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("com.vivo.sprinbootpoi.mapper")
public class SprinbootPoiApplication {

    public static void main(String[] args) {
        SpringApplication.run(SprinbootPoiApplication.class, args);
    }

}
