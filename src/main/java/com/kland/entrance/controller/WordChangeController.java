package com.kland.entrance.controller;

import com.kland.common.config.ServerConfig;
import com.kland.common.entity.CustomGlobalBean;
import com.kland.entrance.service.IWordChangeService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiImplicitParam;
import io.swagger.annotations.ApiImplicitParams;
import io.swagger.annotations.ApiOperation;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

@Api(value = "WordController 测试swagger注解的控制器")
@RestController
@Slf4j
public class WordChangeController {
    @Autowired
    private ServerConfig serverConfig;
    @Autowired
    private IWordChangeService wordChangeService;


    /**
     * Aspose Word html转换word
     * @param inputFilePath
     * @param outFilePath
     */
    @RequestMapping(value ="/changeHtmlToWord", method= RequestMethod.GET)
    @ApiOperation(value = "使用: Aspose Word  Html转化成 Word")
    @ApiImplicitParams({
            @ApiImplicitParam(name="inputFilePath",value = "文件输入路径",paramType = "query",required = true),
            @ApiImplicitParam(name="outFilePath",value = "文件输出位置",paramType = "query")
    })
    public void changeHtmlToWord(String inputFilePath, String outFilePath){
        CustomGlobalBean globalBean = new CustomGlobalBean();
        log.info("全局属性[APPDomain]: {}" + globalBean.getAppDomain());
        log.info("全局属性[WordBasePath]: {}" + globalBean.getWordBasePath());

        log.info("获取服务器路径: " + serverConfig.getUrl());
        log.info("文件输入路径" + inputFilePath);
        log.info("文件输出路径:" + outFilePath);
        wordChangeService.changeHtmlToWord(inputFilePath,outFilePath);
    }


    /*
    @RequestMapping(value ="/changeHtmlToPdf", method= RequestMethod.GET)
    @ApiOperation(value = "使用: PD4ML Html转化成 Word")
    @ApiImplicitParams({
            @ApiImplicitParam(name="inputFilePath",value = "文件输入路径",paramType = "query",required = true),
            @ApiImplicitParam(name="outFilePath",value = "文件输出位置",paramType = "query"),
            @ApiImplicitParam(name="showType",value = "显示类型",paramType = "query",required = true,defaultValue ="0")
    })
    public void changeHtmlToPdf(String inputFilePath, String outFilePath,Integer showType){
        wordChangeService.changeHtmlToPdf(inputFilePath,outFilePath,showType);
    }
    */
}