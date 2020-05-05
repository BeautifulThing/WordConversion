
package com.kland.entrance.service.impl;

import com.kland.common.config.CustomGlobal;
import com.kland.common.util.AsposeUtils;
import com.kland.entrance.service.IWordChangeService;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.zefer.pd4ml.PD4Constants;
import org.zefer.pd4ml.PD4ML;
import org.zefer.pd4ml.PD4PageMark;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

@Slf4j
@Service
public class WordServiceImpl implements IWordChangeService {
    @Autowired
    CustomGlobal customGlobal;
    @Override
    public boolean changeHtmlToWord(String sourcePath, String targetPath) {
        File file = new File(sourcePath);
        if (file.isDirectory()) {
            String [] fileList = file.list();
            if(2 > fileList.length){
                for (String fileName : fileList) {
                    if(fileName.length() > 5){
                        if((".doc") == (fileName.subSequence(fileName.length() - 4, fileName.length()))){
                            try {
                                return AsposeUtils.htmlToDoc(sourcePath, targetPath,customGlobal.getCustomGlobalBean());
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
            return AsposeUtils.htmlToDoc(sourcePath,targetPath,customGlobal.getCustomGlobalBean());
        }
        return false;
    }

    @Override
    public boolean changeHtmlToPdf(String inputFilePath, String outFilePath, Integer showType) {
        File downloadFile = new File(inputFilePath);
        if (!downloadFile.exists()) {
            return false;
        }
        PD4ML pd4ml = new PD4ML();
        PD4PageMark footer = new PD4PageMark();
        footer.setAreaHeight(-1);
        pd4ml.setPageFooter(footer);
        pd4ml.setPageInsets(new Insets(20, 10, 10, 10));
        pd4ml.setHtmlWidth(980);
        pd4ml.setPageSize(PD4Constants.A4);
        try {
            pd4ml.useTTF("java:fonts", true);
        } catch (FileNotFoundException e) {
            log.error("未找到转换pdf的字体jar包");
            return false;
        }
        if (showType == 1) {
            pd4ml.setDefaultTTFs("Arial", "Arial", "Arial");
        } else {
            pd4ml.setDefaultTTFs("Msyh", "Msyh", "Msyh");
        }
        File destFile = new File(outFilePath);
        FileOutputStream fos = null;
        try {
            if(!destFile.exists()){
                destFile.createNewFile();
            }
            fos = FileUtils.openOutputStream(destFile);
            pd4ml.render("file:" + inputFilePath, fos);
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if(null != fos){
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }
}