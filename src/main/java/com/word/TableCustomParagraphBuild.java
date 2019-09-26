package com.word;

import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.math.BigInteger;
import java.util.List;
import java.util.Map;

/**
 * 不存在表格模型，在模板中存在 ##{foreachTable.**}## 的标记
 * 使用默认的表格样式
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/18 16:38
 */
public abstract class TableCustomParagraphBuild extends TableMarge implements ParagraphBuild {

    @Override
    public void replaceParagraph(XWPFDocument document) {

        Map<String, List<String[]>> paramData = getParamData();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        XWPFParagraph paragraph = null;
        while((paragraph = getTableCustom(document.getParagraphs(), paramData)) != null ){
            String tableKey = paragraph.getText();
            XmlCursor cursor = paragraph.getCTP().newCursor();
            XWPFTable table = document.insertNewTbl(cursor);
            document.removeBodyElement(document.getPosOfParagraph(paragraph));
            Map<String, List<MargeTable>> margeData = getMargeData();
            ParagraphBuildHandle.insertTable(table, paramData.get(tableKey), margeData.get(tableKey));
            addTableStyle(table);
        }
    }

    private XWPFParagraph getTableCustom(List<XWPFParagraph> paragraphs, Map<String, List<String[]>> paramData){
        for (XWPFParagraph paragraph : paragraphs) {

            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            String tableKey = ParagraphBuildHandle.checkTableCustom(text);
            if (StringUtils.isNotBlank(tableKey) && paramData.containsKey(tableKey)) {
                ParagraphBuildHandle.replaceParagraph(paragraph, tableKey, true);
                return paragraph;
            }
        }
        return null;
    }




    /**
     * 默认下左右实线黑色边框
     * @param table
     */
    public void addTableStyle(XWPFTable table){
        CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
        CTBorder hBorder = borders.addNewInsideH();
        // 线条类型
        hBorder.setVal(STBorder.Enum.forString("single"));
        // 线条大小
        hBorder.setSz(new BigInteger("1"));
        // 设置颜色
        hBorder.setColor("000000");

        CTBorder vBorder = borders.addNewInsideV();
        vBorder.setVal(STBorder.Enum.forString("single"));
        vBorder.setSz(new BigInteger("1"));
        vBorder.setColor("000000");

        CTBorder lBorder = borders.addNewLeft();
        lBorder.setVal(STBorder.Enum.forString("single"));
        lBorder.setSz(new BigInteger("1"));
        lBorder.setColor("000000");

        CTBorder rBorder = borders.addNewRight();
        rBorder.setVal(STBorder.Enum.forString("single"));
        rBorder.setSz(new BigInteger("1"));
        rBorder.setColor("000000");

        CTBorder tBorder = borders.addNewTop();
        tBorder.setVal(STBorder.Enum.forString("single"));
        tBorder.setSz(new BigInteger("1"));
        tBorder.setColor("000000");

        CTBorder bBorder = borders.addNewBottom();
        bBorder.setVal(STBorder.Enum.forString("single"));
        bBorder.setSz(new BigInteger("1"));
        bBorder.setColor("000000");
        table.getCTTbl().getTblPr().getTblW().setType(STTblWidth.DXA);
        table.getCTTbl().getTblPr().getTblW().setW(new BigInteger("8300"));

    }

    /**
     *
     * 注入业务数据
     * @return
     */
    public abstract Map<String, List<String[]>> getParamData();


}
