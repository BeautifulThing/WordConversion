package com.lin.entrance.service;


public interface IWordChangeService {
    boolean changeHtmlToWord(String inputFilePath, String outFilePath);

    boolean changeHtmlToPdf(String inputFilePath, String outFilePath, Integer showType);

    boolean changePdfToWord(String inputFilePath, String outFilePath) throws Exception;
}
