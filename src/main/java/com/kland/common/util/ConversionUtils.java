package com.kland.common.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.fonts.BestMatchingMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.PageSizePaper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.IOException;
import java.util.Date;

@Slf4j
public class ConversionUtils {
    public static boolean conversionSimpleFile(String srcFilePath, String outFilePath){
        log.info("######## ======== 开始执行转换...");
        log.info("srcFilePath: {}",srcFilePath);
        log.info("outFilePath: {}",outFilePath);
        // File fileIn = new File(srcFilePath);
        // System.out.println("file Exist :" + fileIn.exists());
        // 1.判断文件是否存在
        // if(!fileIn.exists()){
        //    return false;
        // }
        // 2.判断outFilePath 是否为Null
        if(StringUtils.isNoneBlank(outFilePath)){
            return executeConversion(srcFilePath,outFilePath);
        }else{
            // 3.如果 outFilePath 为Null,程序自动生成输出路径地址
            outFilePath = resultOutFilePath();
            // 4.采用 Docx4j 转换html至doc
            return executeConversion(srcFilePath,outFilePath);
        }
    }


    private static boolean executeConversion(String srcFilePath, String outFilePath) {
        // 1.将.Html页面转化成org.jsoup.nodes.Document 对象
        boolean flag = false;
        Document document = null;
        WordprocessingMLPackage wordMLPackage = null;
        try {
            long old = System.currentTimeMillis();
            File input  = new File(srcFilePath);
            // 2.采取Jsoup 处理html内容(xhtml)
            // document = handlerHtmlContent(Jsoup.parse(input, "UTF-8"));
            document = handlerHtmlContent(Jsoup.connect(srcFilePath).get());
            // 3.将xhtml 转换成 word
            wordMLPackage = xhtmlToWord(document);
            File out = new File(outFilePath);
            wordMLPackage.save(out);
            long now = System.currentTimeMillis();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒");  //转化用时
            flag = true;
            if (log.isDebugEnabled()) {
                log.debug("Save to [.doc]: {}", out.getAbsolutePath());
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }

    public static Document handlerHtmlContent(Document document) {
        Elements trs = document.getElementsByTag("tr");
        for (Element tr : trs) {
            System.out.println(tr);
        }
        


        if (log.isDebugEnabled()) {
            log.debug("baseUri: {}", document.baseUri());
        }
        // 去除Script标签
        for (Element scripts : document.getElementsByTag("script")) {
            scripts.remove();
        }
        // 去除A 标签 href 属性 和 onclick 点击事件
        /*
        for (Element as : document.getElementsByTag("a")){
            as.removeAttr("onclick");
            as.removeAttr("href");
        }
        */
        // 将link中的地址替换为绝对地址
        Elements links = document.getElementsByTag("link");
        for (Element element : links) {
            String href = element.absUrl("href");
            if (log.isDebugEnabled()) {
                log.debug("href: {} -> {}", element.attr("href"), href);
            }
            element.attr("href", href);
        }

        
        document.outputSettings()
                .syntax(Document.OutputSettings.Syntax.xml)
                .escapeMode(Entities.EscapeMode.xhtml);
        if (log.isDebugEnabled()) {
            log.info("document.html: " + document.html());
        }
        return document;
    }

    static WordprocessingMLPackage xhtmlToWord(Document document) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage(PageSizePaper.valueOf("A4"), true);
        // settingFont(wordMLPackage);
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        // mainDocumentPart.addObject();

        ObjectFactory factory = Context.getWmlObjectFactory();
        RPr rpr = factory.createRPr();
        RFonts font = new RFonts();
        font.setAscii("宋体");
        font.setEastAsia("宋体");//经测试发现这个设置生效
        rpr.setRFonts(font);

        XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordMLPackage);
        // xhtmlImporter.setHyperlinkStyle("Hyperlink");
        wordMLPackage.getMainDocumentPart().getContent().addAll(
                xhtmlImporter.convert(document.html(), document.baseUri()));
        return wordMLPackage;
    }

    static void settingFont(WordprocessingMLPackage wordMLPackage) throws Exception {
        Mapper fontMapper  = new BestMatchingMapper();
        fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
        wordMLPackage.setFontMapper(fontMapper);
        String fontFamily = "宋体";
        RFonts fonts = Context.getWmlObjectFactory().createRFonts();
        fonts.setAsciiTheme(null);
        fonts.setAscii(fontFamily);
        wordMLPackage.getMainDocumentPart()
                .getPropertyResolver()
                .getDocumentDefaultRPr()
                .setRFonts(fonts);
    }

    static String resultOutFilePath () {
        return "D:/OutFilePath/auto-" + new Date().getTime() + "-" + ((int) Math.random() * 100000) + ".doc";
    }
}