package com.kland.common.util;


import com.aspose.words.*;
import org.apache.commons.lang3.StringUtils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

public class AsposeUtils {
    /**
     * Aspose
     * @return
     */
    public static boolean getLicense() {
        boolean flag = false;
        License asposeLicense = null;
        InputStream is =  Thread.currentThread().getContextClassLoader().getResourceAsStream("license.xml");
        try {
            asposeLicense = new License();
            asposeLicense.setLicense(is);
            flag = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }

    public static boolean changeSimpleFile(String srcFilePath, String outFilePath){

        // File fileIn = new File(srcFilePath);
        // System.out.println("file Exist :" + fileIn.exists());
        // 1.判断文件是否存在
        // if(!fileIn.exists()){
        //    return false;
        // }
        // 2.判断outFilePath 是否为Null
        if(StringUtils.isNoneBlank(outFilePath)){
            return executeChange(srcFilePath,outFilePath);
        }else{
            // 3.如果 outFilePath 为Null,程序自动生成输出路径地址
            outFilePath = resultOutFilePath();
            return executeChange(srcFilePath,outFilePath);
        }


    }

    private static boolean executeChange(String srcFilePath, String outFilePath) {
        if(!getLicense()){
            return false;
        }
        // File file = new File(srcFilePath);
        OutputStream os = null;
        Document document = null;

        DocumentBuilder builder = null;
        try {
            long old = System.currentTimeMillis();
            document = new Document(srcFilePath);
            builder = new DocumentBuilder(document);

            Font font = builder.getFont();
            font.setSize(22);
            font.setNameFarEast("宋体");

            ParagraphFormat paragraphFormat = builder.getParagraphFormat();
            paragraphFormat.setLineSpacing(12);

            PageSetup pageSetup = builder.getPageSetup();
            pageSetup.setPaperSize(PaperSize.A4);
            pageSetup.setVerticalAlignment(PageVerticalAlignment.TOP);
            builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);


            os = new FileOutputStream(outFilePath);
            document.save(os, SaveFormat.DOC);

            long now = System.currentTimeMillis();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒");  //转化用时
            os.close();
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    static String resultOutFilePath () {
        return "D:/OutFilePath/auto-" + new Date().getTime() + "-" + ((int) Math.random() * 100000) + ".doc";
    }
}