
package com.lin.entrance.service.impl;


import com.lin.common.config.CustomGlobal;
import com.lin.common.util.AsposePdfUtils;
import com.lin.common.util.AsposeWordUtils;
import com.lin.entrance.service.IWordChangeService;
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
import java.net.MalformedURLException;

@Slf4j
@Service
public class WordServiceImpl implements IWordChangeService {
    @Autowired
    CustomGlobal customGlobal;
    @Override
    public boolean changeHtmlToWord(String sourcePath, String targetPath) {
        return AsposeWordUtils.htmlToDoc(sourcePath,targetPath,customGlobal.getCustomGlobalBean());
    }

    @Override
    public boolean changeHtmlToPdf(String inputFilePath, String outFilePath, Integer showType) {
        File downloadFile = new File(inputFilePath);
        if (!downloadFile.exists()) {
            return false;
        }
        PD4ML pd4ml = new PD4ML();
        PD4PageMark footer = new PD4PageMark();
        footer.setHtmlTemplate("<p style='font-size:16px;float:right;'><span>page $[page] of $[total]</span><p>");
        footer.setAreaHeight(-1);
        pd4ml.setPageFooter(footer);
        pd4ml.setPageInsets(new Insets(20, 10, 10, 10));
        pd4ml.setHtmlWidth(980);
        pd4ml.setPageSize(PD4Constants.A4);
        File destFile = null;
        FileOutputStream fos = null;
        try {
            pd4ml.useTTF("java:fonts", true);
            if (showType == 1) {
                pd4ml.setDefaultTTFs("Arial", "Arial", "Arial");
            } else {
                pd4ml.setDefaultTTFs("Msyh", "Msyh", "Msyh");
            }
            destFile = new File(outFilePath);
            if(!destFile.exists()){
                destFile.createNewFile();
            }
            fos = FileUtils.openOutputStream(destFile);
            pd4ml.render("file:" + inputFilePath, fos);
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (MalformedURLException e) {
            e.printStackTrace();
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
        return false;
    }

    @Override
    public boolean changePdfToWord(String inputFilePath, String outFilePath) throws Exception {
       return AsposePdfUtils.pdfToWord(inputFilePath,outFilePath);
    }

}