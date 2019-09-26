package com.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.List;
import java.util.Map;

/**
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/17 17:25
 */
public abstract class TextParagraphBuild implements ParagraphBuild {

    @Override
    public void replaceParagraph(XWPFDocument document) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            ParagraphBuildHandle.replaceParagraph(getParamData(), paragraph, false);
        }
    }



    /**
     *
     * 注入业务数据
     * @return
     */
    public abstract Map<String, String> getParamData();
}
