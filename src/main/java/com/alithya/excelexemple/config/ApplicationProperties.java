package com.alithya.excelexemple.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Data
@Component
@ConfigurationProperties
public class ApplicationProperties {

    private String excelFileName;

}
