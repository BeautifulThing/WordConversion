package com.kland.common.util;


import com.aspose.words.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.Date;

@Slf4j
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
            long start = System.currentTimeMillis();
            String content = FileUtils.readFileToString(new File(srcFilePath), "UTF-8");
            if(content.contains("logo1_en")){
                isEnglish = true;
            }
            content = content.replaceAll("<xml><w:WordDocument><w:View>Normal</w:View></w:WordDocument></xml>", "");
            content = content.replaceAll("<img[^<]*?logo1[^<]*?/>", "");
            // log.info("文件内容信息: {}" ,content);
            document = new Document();
            builder = new DocumentBuilder(document);
            builder.insertHtml(content);

            PageSetup pageSetup = builder.getPageSetup();
            pageSetup.setPaperSize(PaperSize.A4);
            pageSetup.setOrientation(Orientation.PORTRAIT);
            pageSetup.setVerticalAlignment(PageVerticalAlignment.CENTER);
            // 72 90

            ParagraphCollection paragraphs = document.getFirstSection().getBody().getParagraphs();
            for (Paragraph paragraph : paragraphs) {
                ParagraphFormat paragraphFormat = paragraph.getParagraphFormat();
                // paragraphFormat.setFirstLineIndent(24.0d);
                // paragraphFormat.setLineSpacing(18.0d);
                // paragraphFormat.setLineSpacingRule(2);
                // paragraphFormat.setFirstLineIndent(24);
                paragraphFormat.setLeftIndent(0);
                paragraphFormat.setSpaceBeforeAuto(true);
                paragraphFormat.setSpaceAfterAuto(true);
            }

            TableCollection tables = document.getFirstSection().getBody().getTables();
            for (Table table : tables) {
                RowCollection rows = table.getRows();
                for (Row row : rows) {
                    CellCollection cells = row.getCells();
                    for (Cell cell : cells) {
                        cell.getParagraphs().forEach(paragraph -> {
                            ParagraphFormat paragraphFormat = paragraph.getParagraphFormat();
                            paragraphFormat.setSpaceBeforeAuto(true);
                            paragraphFormat.setSpaceAfterAuto(true);
                        });
                    }
                }
            }
            /* 设置页眉**/
            builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
            // builder.getCurrentParagraph().getParagraphFormat().getBorders().getBottom().setLineStyle(LineStyle.SINGLE);
            // builder.getCurrentParagraph().getParagraphFormat().getBorders().getBottom().setLineWidth(1.0);
            builder.getCurrentParagraph().getParagraphFormat().getTabStops().add(410,TabAlignment.RIGHT,TabLeader.NONE);
            File imgFile = null;
            if(isEnglish){
                imgFile = new File("src/main/resources/static/img/logo1_en.png");
            }else{
                imgFile = new File("src/main/resources/static/img/logo1_zh_taa.png");
            }
            InputStream is=new FileInputStream(imgFile);
            builder.insertImage(is,286,18);
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
            long end = System.currentTimeMillis();
            log.info("执行转换所花时间: {} 秒!",(start - end) % 60);
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