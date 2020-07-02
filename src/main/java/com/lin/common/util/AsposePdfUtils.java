package com.lin.common.util;

import com.aspose.pdf.Document;
import com.aspose.pdf.HtmlLoadOptions;
import com.aspose.pdf.License;
import com.aspose.pdf.SaveFormat;

import java.io.*;


public class AsposePdfUtils {
    /**
     *获取aspose pdf license
     * @return
     */
    public static boolean getWordLicense(){
        boolean result = false;
        try {
            InputStream is =Thread.currentThread().getContextClassLoader().getResourceAsStream("license-pdf.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            result=false;
        }
        return result;
    }


    public static boolean pdfToWord(String srcFilePath, String outFilePath) throws Exception {
        if(!getWordLicense()){
            throw new Exception("aspose.pdf license错误");
        }
        InputStream in = null;
        OutputStream os = null;
        File outFile = null;
        try {
            in = new FileInputStream(srcFilePath);
            outFile = new File(outFilePath);
            if(!outFile.exists()){
                outFile.createNewFile();
            }
            os = new FileOutputStream(outFilePath);
            Document pdfDoc = new Document(in);
            pdfDoc.save(os, SaveFormat.Doc);
            os.flush();
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if(null != os){
                os.close();
            }
        }
        return false;
    }
}
