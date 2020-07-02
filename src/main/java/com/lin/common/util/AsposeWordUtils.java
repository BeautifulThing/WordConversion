package com.lin.common.util;


import com.aspose.words.*;
import com.lin.common.entity.CustomGlobalBean;
import com.lin.common.other.ReplaceAndInsertImage;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class AsposeWordUtils {
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

    public static boolean htmlToDoc(String srcFilePath, String outFilePath, CustomGlobalBean customGlobalBean){
        File fileIn = new File(srcFilePath);
        if(!fileIn.exists()){
            return false;
        }
        if(StringUtils.isNoneBlank(outFilePath)){
            return executeChange(srcFilePath,outFilePath,customGlobalBean);
        }else{
            outFilePath = resultOutFilePath();
            return executeChange(srcFilePath,outFilePath,customGlobalBean);
        }
    }

    private static boolean executeChange(String srcFilePath, String outFilePath,CustomGlobalBean customGlobalBean) {
        if(!getLicense()){
            return false;
        }
        OutputStream os = null;
        Document document = null;
        DocumentBuilder builder = null;
        boolean isEnglish = false;
        Map<String,Object> pictureListMap = new LinkedHashMap<>();
        String logoName = "";
        try {
            long start = System.currentTimeMillis();
            String content = FileUtils.readFileToString(new File(srcFilePath), "UTF-8");
            if(content.contains("logo1_en")){
                isEnglish = true;
            }
            content = content.replaceAll("<xml><w:WordDocument><w:View>Normal</w:View></w:WordDocument></xml>", "");
            content = content.replaceAll("<img[^<]*?logo1[^<]*?/>", "");
            content = handlePicture(pictureListMap,content,"",customGlobalBean);

            document = new Document();
            builder = new DocumentBuilder(document);
            builder.insertHtml(content);

            PageSetup pageSetup = builder.getPageSetup();
            pageSetup.setPaperSize(PaperSize.A4);
            pageSetup.setOrientation(Orientation.PORTRAIT);
            pageSetup.setVerticalAlignment(PageVerticalAlignment.CENTER);
            pageSetup.setTopMargin(72);
            pageSetup.setBottomMargin(72);
            pageSetup.setLeftMargin(54);
            pageSetup.setRightMargin(54);
            pageSetup.setHeaderDistance(40);
            handleParagraph(document);
            settingPageValue(builder,isEnglish,pictureListMap,null,content);
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
            log.info("共耗时: {} 秒!",(start - end) /  1000.0);
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

    /**
     * 处理段落
     *  1.默认转化段落之间没有上下间距
     *  2.Table标签里面也包含段落(对应也处理)
     * @param document
     */
    static void handleParagraph(Document document){
        ParagraphCollection paragraphs = document.getFirstSection().getBody().getParagraphs();
        for (Paragraph paragraph : paragraphs) {
            ParagraphFormat paragraphFormat = paragraph.getParagraphFormat();
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
    }

    /**
     *
     * @param content
     * @param dirPrefix
     * @param customGlobalBean
     * @return
     */
    static String handlePicture(Map<String,Object> pictureListMap,String content,String dirPrefix,CustomGlobalBean customGlobalBean){
        if(StringUtils.isBlank(content)){
            return "";
        }
        Matcher matcher = Pattern.compile("(?i)<img[^<]+src\\s*=\\s*['\"]([^'\">]+)['\"][^>]*>").matcher(content);
        Map<String,Object> imgMap = new LinkedHashMap<>();
        Map<String,Object> imgAttrMap = new LinkedHashMap();
        Integer i = 0;
        String path = "";
        while (matcher.find()) {
            String imgSrcAddress = dirPrefix + matcher.group(1);

            /*
            Matcher heigthMatcher = Pattern.compile("height[\\s]*:[\\s]*(\\d+)(px)").matcher(matcher.group(0));
            while (heigthMatcher.find()) {
                imgAttrMap.put("height",Integer.parseInt(heigthMatcher.group(1)) / 2);
            }
            Matcher widthMatcher = Pattern.compile("width[\\s]*:[\\s]*(\\d+)(px)").matcher(matcher.group(0));
            while (heigthMatcher.find()) {
                imgAttrMap.put("width",Integer.parseInt(widthMatcher.group(1)) / 2);
            }
            */

            path = imgSrcAddress;
            if(imgSrcAddress.startsWith("/") && !imgSrcAddress.contains(customGlobalBean.getAppDomain())){
                path = customGlobalBean.getAppDomain() + imgSrcAddress;
            }else {
                String replaceStr = "http[s]?://" + customGlobalBean.getAppDomain();
                if (imgSrcAddress.contains(":80")) {
                    replaceStr += ":80";
                }
                path = imgSrcAddress.replaceAll(replaceStr, customGlobalBean.getWordBasePath());
            }
            imgAttrMap.put("path",path);
            pictureListMap.put("#*#*"+i+"#*#*",imgAttrMap);
            imgMap.put("#*#*" + i +"#*#*",matcher.group(0));
            i++;
        }
        for (String key : imgMap.keySet()) {
            content = content.replaceFirst(imgMap.get(key).toString().replace("(","\\(").replace(")","\\)"),key);
        }
        return content;
    }

    /**
     * 设置页眉页码(同时替换中英文LOGO)
     * @param builder
     * @param isEnglish
     * @throws Exception
     */
    static void settingPageValue(DocumentBuilder builder, boolean isEnglish, Map<String,Object> pictureListMap,String logoName,String content) throws Exception {
        /* 设置页眉**/
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getCurrentParagraph().getParagraphFormat().getTabStops().add(410,TabAlignment.RIGHT,TabLeader.NONE);
        File imgFile = null;
        if(isEnglish){
            imgFile = new File("src/main/resources/static/img/logo1_en.png");
        }else{
            imgFile = new File("src/main/resources/static/img/logo1_zh.png");
        }
        builder.insertImage(new FileInputStream(imgFile),286,18);
        builder.write(ControlChar.TAB);

        /* 设置页脚-页码**/
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.insertField("PAGE", "");
        builder.write("/");
        builder.insertField("NUMPAGES", "");


        /* 替换图片**/
        for (String key: pictureListMap.keySet()) {
            Map value = (Map)pictureListMap.get(key);
            replaceImage(key,value,builder,content);
        }
    }

    private static boolean replaceImage(String key,Map image, DocumentBuilder builder,String content) {
        Document document = builder.getDocument();
        if(StringUtils.isNoneBlank(key)){
            Integer tagNumber = Integer.valueOf(key.replaceAll("[#*]+", ""));
            String regex = String.format("[#*]+%d[#*]+", tagNumber);
            File imgFile = new File((String)image.get("path"));
            InputStream is = null;
            try {
                is = new FileInputStream(imgFile);
                if (imgFile.exists()) {
                    // 替换文档中Img(本地图片文件来源配置路径)
                    document.getRange().replace(Pattern.compile(regex), new ReplaceAndInsertImage((String) image.get("path")), false);
                }
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            }finally {
                if(null != is){
                    try {
                        is.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        return true;
    }

    static String resultOutFilePath () {
        return "D:/OutFilePath/auto-" + new Date().getTime() + "-" + ((int) Math.random() * 100000) + ".doc";
    }
}