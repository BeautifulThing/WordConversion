package com.lin.common.config;

import lombok.extern.slf4j.Slf4j;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import springfox.documentation.builders.ApiInfoBuilder;
import springfox.documentation.builders.PathSelectors;
import springfox.documentation.builders.RequestHandlerSelectors;
import springfox.documentation.service.ApiInfo;
import springfox.documentation.service.Contact;
import springfox.documentation.spi.DocumentationType;
import springfox.documentation.spring.web.plugins.Docket;
import springfox.documentation.swagger2.annotations.EnableSwagger2;

@Configuration
@EnableSwagger2
@Slf4j
public class SwaggerAppConfig {
    @Bean
    public Docket createRestApi() {
        return new Docket(DocumentationType.SWAGGER_2)
                .apiInfo(getApiInfo())
                .select()
                .apis(RequestHandlerSelectors.basePackage("com.kland.entrance"))
                .paths(PathSelectors.any())
                .build();
    }

    public ApiInfo getApiInfo() {
        return new ApiInfoBuilder()
                //　页面标题
                .title("Spring Boot 使用 Swagger2 构建RESTful API")
                //　创建人
                .contact(new Contact("Trey", "", ""))
                //　版本号
                .version("1.0")
                //　描述
                .description("API 描述")
                .build();
    }
}