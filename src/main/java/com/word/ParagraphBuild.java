package com.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * 段落构建器
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/17 17:21
 */
public interface ParagraphBuild {

    /**
     * 段落替换
     * @param document
     */
    public void replaceParagraph(XWPFDocument document);
}
