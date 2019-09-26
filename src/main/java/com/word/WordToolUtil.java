package com.word;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/17 17:04
 */
@Slf4j
public class WordToolUtil{


    public static boolean handleWordTemplate(String inPath, String outPath, Business ... businessArr){

        try {
            if (businessArr.length == 0){
                throw new WordException("Business can not be null");
            }
            File outFile = new File(outPath);
            if (!outFile.exists()){
                outFile.createNewFile();
            }else if(!(outFile.isFile() && outFile.canWrite())){
                throw new WordException("outPath File can not be create");
            }
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inPath));
            BusinessBuildRegister register = new BusinessBuildRegister();
            register.register(businessArr);
            register.replaceWord(document);
            FileOutputStream stream = new FileOutputStream(outFile);
            document.write(stream);
            stream.close();
            return true;

        } catch (IOException e) {
            e.printStackTrace();
            log.error("",e);
        } catch (WordException e) {
            e.printStackTrace();
            log.error("",e);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            log.error("",e);
        } catch (InstantiationException e) {
            e.printStackTrace();
            log.error("",e);
        }
        return false;
    }
}
