package com.kland.entrance.service;


public interface IWordChangeService {
    boolean conversionHtmlToWord(String inPath,String outPath);

    boolean changeHtmlToWord(String inputFilePath, String outFilePath);

    boolean changeHtmlToPdf(String inputFilePath, String outFilePath, Integer showType);
}
