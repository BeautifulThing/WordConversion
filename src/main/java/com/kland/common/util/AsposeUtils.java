package com.kland.common.util;


import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.Date;

public class AsposeUtils {
    /**
     * Aspose
     * @return
     */
    public static boolean getLicense() {
        boolean flag = false;
        License asposeLicense = null;
        InputStream is =  Thread.currentThread().getContextClassLoader().getResourceAsStream("license-word.xml");
        try {
            asposeLicense = new License();
            asposeLicense.setLicense(is);
            flag = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }

    public static boolean htmlToDoc(String srcFilePath, String outFilePath){

        File fileIn = new File(srcFilePath);
        if(!fileIn.exists()){
            return false;
        }
        if(StringUtils.isNoneBlank(outFilePath)){
            return executeChange(srcFilePath,outFilePath);
        }else{
            outFilePath = resultOutFilePath();
            return executeChange(srcFilePath,outFilePath);
        }
    }

    private static boolean executeChange(String srcFilePath, String outFilePath) {
        if(!getLicense()){
            return false;
        }
        OutputStream os = null;
        Document document = null;
        DocumentBuilder builder = null;
        boolean isEnglish = false;
        try {
            String content = FileUtils.readFileToString(new File(srcFilePath), "UTF-8");
            if(content.contains("logo1_en")){
                isEnglish = true;
            }
            content = content.replaceAll("<xml><w:WordDocument><w:View>Normal</w:View></w:WordDocument></xml>", "");
            content = content.replaceAll("<img[^<]*?logo1[^<]*?/>", "");
            System.out.println(content);
            document = new Document();
            builder = new DocumentBuilder(document);
            builder.insertHtml(content);
            /* 设置页眉**/
            builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
            builder.getParagraphFormat().getTabStops().add(410,TabAlignment.RIGHT,TabLeader.NONE);
            File file = new File("src/main/resources/static/img/logo1_zh_taa.png");
            InputStream is=new FileInputStream(file);
            builder.insertImage(is,148,15);
            builder.write(ControlChar.TAB);
            /* 设置页脚-页码**/
            builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
            builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
            builder.insertField("PAGE", "");
            builder.write("/");
            builder.insertField("NUMPAGES", "");

            File targetFile = new File(outFilePath);
            if(targetFile.exists()) {
                // 如果存在就执行删除操作
                targetFile.delete();
            }
            // 删除完毕后执行创建
            targetFile.createNewFile();
            // IO封装输出流
            os = new FileOutputStream(targetFile);
            document.save(os, SaveFormat.DOC);
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if(null != os) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return false;
    }

    static String resultOutFilePath () {
        return "D:/OutFilePath/auto-" + new Date().getTime() + "-" + ((int) Math.random() * 100000) + ".doc";
    }
}