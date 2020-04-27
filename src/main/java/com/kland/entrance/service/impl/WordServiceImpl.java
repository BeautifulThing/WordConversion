
package com.kland.entrance.service.impl;

import com.kland.common.util.AsposeUtils;
import com.kland.common.util.ConversionUtils;
import com.kland.entrance.service.IWordService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.Arrays;
import java.util.stream.Stream;

@Slf4j
@Service
public class WordServiceImpl implements IWordService {
    @Override
    public boolean conversionHtmlToWord(String sourcePath, String targetPath) {
        /*
        // 1.File IO 对象封装路径
        File file = new File(sourcePath);
        // 2.判断是不是文件夹
        if (file.isDirectory()) {
            log.info("##############------ 文件夹名称为: {}",sourcePath);
            // 3.遍历文件夹
            String [] fileList = file.list();
            Stream<String> stream = Arrays.stream(fileList);
            // 4. 注意: 2 篇以上法规不进行转换
            if(2 > fileList.length){
                stream.forEach(fileName ->{
                    if(fileName.length() > 5){
                        // 5.判断文件名称后缀是否是.html
                        if(".html" == (fileName.subSequence(fileName.length() - 4, fileName.length()))){
                            try {
                                // 6.执行转换逻辑
                                ConversionUtils.conversionSimpleFile(sourcePath, targetPath);
                            } catch (Exception e) {
                                e.printStackTrace();
                                log.error("conversion error "+ e.getMessage(),e);
                            }
                        }
                    }
                });
            }else{
                log.info("more than {} files",fileList.length);
            }
        }else{
            return ConversionUtils.conversionSimpleFile(sourcePath ,targetPath);
        }
        return true;
        */
        ConversionUtils.conversionSimpleFile(sourcePath, targetPath);
        return true;
    }

    @Override
    public boolean changeHtmlToWord(String sourcePath, String targetPath) {
        File file = new File(sourcePath);
        if (file.isDirectory()) {
            String [] fileList = file.list();
            Stream<String> stream = Arrays.stream(fileList);
            if(2 > fileList.length){
                for (String fileName : fileList) {
                    if(fileName.length() > 5){
                        if(".doc" == (fileName.subSequence(fileName.length() - 4, fileName.length()))){
                            try {
                                return AsposeUtils.htmlToDoc(sourcePath, targetPath);
                            } catch (Exception e) {
                                e.printStackTrace();
                                log.error("conversion error "+ e.getMessage(),e);
                            }
                        }
                    }
                }
            }else{
                log.info("more than {} files",fileList.length);
            }
        }else{
            return AsposeUtils.htmlToDoc(sourcePath,targetPath);
        }
        return false;
    }
}