package com.lin.common.config;

import com.lin.common.entity.CustomGlobalBean;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;

@Component
public class CustomGlobal {
    @Value("${app.domain}")
    private String appDomain;
    @Value("${word.base.path}")
    private String wordBasePath;

    @Bean("customGlobalBean")
    public CustomGlobalBean getCustomGlobalBean(){
        CustomGlobalBean customGlobalBean = new CustomGlobalBean();
        customGlobalBean.setAppDomain(appDomain);
        customGlobalBean.setWordBasePath(wordBasePath);
        return customGlobalBean;
    }
}
