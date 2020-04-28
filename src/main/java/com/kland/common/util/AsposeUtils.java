package com.kland.common.util;


import com.aspose.words.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.asm.FieldVisitor;

import java.io.*;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
            String content = FileUtils.readFileToString(new File(srcFilePath), "UTF-8");
            if(content.contains("logo1_en")){
                isEnglish = true;
            }
            content = content.replaceAll("<xml><w:WordDocument><w:View>Normal</w:View></w:WordDocument></xml>", "");
            content = content.replaceAll("<img[^<]*?logo1[^<]*?/>", "");
            // content = content.replaceAll("<img[^<]*?qrCode[^<]*?/>", "");
            log.info("文件内容信息: {}" ,content);


            /** 正则匹配*/
            String regex ="/(?i)<img[^<]+src\\s*=\\s*['\"]([^'\">]+)['\"][^>]*>/";
            Pattern compile = Pattern.compile(regex);
            Matcher matcher = compile.matcher(content);
            // System.out.println(matcher.find());
            if (matcher.find()) {
                System.out.println("Found value: " + matcher.group(1) );
            }

            document = new Document();
            builder = new DocumentBuilder(document);
            builder.insertHtml(content);


            // document.getRange().replace(ControlChar.SPACE_CHAR, ControlChar.LF);
            PageSetup pageSetup = builder.getPageSetup();
            pageSetup.setPaperSize(PaperSize.A4);
            pageSetup.setOrientation(Orientation.PORTRAIT);
            pageSetup.setVerticalAlignment(PageVerticalAlignment.TOP);
            pageSetup.setLeftMargin(42);
            pageSetup.setRightMargin(42);


            document.getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true).getCount();

            ParagraphCollection paragraphs = document.getFirstSection().getBody().getParagraphs();
            for (Paragraph paragraph : paragraphs) {
                // builder.write(ControlChar.LF);
                
                ParagraphFormat paragraphFormat = paragraph.getParagraphFormat();
                // paragraphFormat.setFirstLineIndent(24.0d);
                // paragraphFormat.setLineSpacing(18.0d);
                // paragraphFormat.setLineSpacingRule(2);

                paragraphFormat.setFirstLineIndent(24);
                paragraphFormat.setLineSpacing(1.0);
            }


            TableCollection tables = document.getFirstSection().getBody().getTables();
            builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
            for (Table table : tables) {
                table.setLeftIndent(0.0d);
            }
            
            /* 设置页眉**/
            builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
            builder.getParagraphFormat().getTabStops().add(410,TabAlignment.RIGHT,TabLeader.NONE);
            File imgFile = null;
            if(isEnglish){
                imgFile = new File("src/main/resources/static/img/logo1_en.png");
            }else{
                imgFile = new File("src/main/resources/static/img/logo1_zh_taa.png");
            }
            InputStream is=new FileInputStream(imgFile);
            builder.insertImage(is,328,28);
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