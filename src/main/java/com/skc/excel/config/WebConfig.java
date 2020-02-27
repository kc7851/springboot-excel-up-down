package com.skc.excel.config;

import io.cogi.spring.constant.MediaType;
import io.cogi.spring.servlet.ExcelViewResolver;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.ContentNegotiationConfigurer;
import org.springframework.web.servlet.config.annotation.ViewResolverRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class WebConfig implements WebMvcConfigurer {

    @Bean
    public ExcelViewResolver xlsViewResolver() {
        ExcelViewResolver xlsViewResolver = new ExcelViewResolver();
        xlsViewResolver.setPrefix("/WEB-INF/excel/");

        return xlsViewResolver;
    }

    @Override
    public void configureViewResolvers(ViewResolverRegistry registry) {
        registry.viewResolver(xlsViewResolver());
    }

}
