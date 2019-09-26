package com.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * word 业务构建器
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/17 17:21
 */
@FunctionalInterface
public interface Business {


    /**
     * 注入段落构建器
     * @param document
     * @return
     */
    public ParagraphBuild paragraphBuild(XWPFDocument document);
}
