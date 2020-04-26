package com.kland.entrance.service;


public interface IWordService {
    boolean conversionHtmlToWord(String inPath,String outPath);

    boolean changeHtmlToWord(String inputFilePath, String outFilePath);
}
