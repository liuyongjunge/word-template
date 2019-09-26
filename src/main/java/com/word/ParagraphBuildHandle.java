package com.word;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

/**
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/17 17:21
 */
public class ParagraphBuildHandle<main> {

    public static Pattern p1 = Pattern.compile("##\\{(\\w+)\\}##");
    public static Pattern p2 = Pattern.compile("\\$\\{([\\w|\\-|/]+)\\}\\$");
    public static Pattern p3 = Pattern.compile("##\\{foreachTable\\.(\\w+)\\}##");


    /**
     * 判断段落中是否存在 ##{foreachTable.**}##  标记
     * @param text
     * @return
     */
    public static String checkTableCustom(String text){
        boolean check  =  false;
        Matcher m = p3.matcher(text);
        if (m.find()){
            return m.group(1);
        }
        return "";
    }


    /**
     * 判断表格第一行一列是否存在 ##{**}##  标记
     * @param table
     * @return
     */
    public static String checkTableEach(XWPFTable table){
        List<XWPFTableRow> rows = table.getRows();
        String rowText = rows.get(0).getCell(0).getText();
        Matcher m = p1.matcher(rowText);
        if (m.find()){
            return m.group(1);
        }
        return "";
    }

    /**
     * 判断文本中时候包含${**}$  标记
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    public static boolean checkText(String text){
        boolean check  =  false;
        Matcher m = p2.matcher(text);
        return m.find();
    }


    /**
     * 匹配传入信息集合与模板
     * @param value 模板需要替换的区域
     * @param paramData 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> paramData){
        Matcher m = p2.matcher(value);
        while (m.find()){
            int count = m.groupCount();
            String[] keys = Stream.iterate(0, i->i+1).limit(count).map(i->m.group(i+1)).toArray(String[]::new);
            String[] values = Stream.of(keys).map(key->{
                Object tmp = paramData.get(key);
                if (tmp == null){
                    return "";
                }
                if (tmp instanceof String){
                    return (String)tmp;
                }else {
                    return tmp.toString();
                }
            }).toArray(String[]::new);
            String[] matchers = Stream.of(keys).map(key->String.format("${%s}$", key)).toArray(String[]::new);
            value = StringUtils.replaceEach(value, matchers, values);
        }
        return value;
    }

    /**
     * 遍历表格
     * @param rows 表格行对象
     * @param paramDate 需要替换的信息集合
     */
    public static void eachTableRow(Map<String, String> paramDate, XWPFTableRow ... rows){
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if(checkText(cell.getText())){
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        replaceParagraph(paramDate, paragraph, true);
                    }
                }
            }
        }
    }



    /**
     * 设置Run 文本（文本中保护 \r\n 则换行）
     * @param paragraph
     * @param text
     * @param pos
     */
    public static void insertText(XWPFParagraph paragraph,String text, int pos, boolean isTable){
        String[] textArr = text.split("\\r\\n");
        for (int i = 0; i < textArr.length; i++) {
            XWPFRun newRun = paragraph.insertNewRun(pos+i);
            newRun.setText(textArr[i],0);
            if (i+1 < textArr.length){
                if (isTable){
                    newRun.addBreak();
                }else {
                    newRun.addCarriageReturn();
                }
            }
        }
    }

    /**
     * 循环遍历表格数据
     * @param paramDate
     * @param table
     * @param margeTables
     */
    public static void eachTableData(List<Map<String, String>> paramDate, XWPFTable table, List<TableMarge.MargeTable> margeTables){

        int size = paramDate.size();
        for (int i=0;i<size;i++){
            Map<String, String> dataMap = paramDate.get(i);
            int rows = table.getRows().size();
            XWPFTableRow row = table.getRow(rows -1);
            if( i+1 < size){
                XWPFTableRow newRow = table.insertNewTableRow(rows);
                copyRow(row, newRow);
            }
            eachTableRow(dataMap, row);
        }
        table.removeRow(0);
        mergeTable(table, margeTables);
    }


    /**
     * 拷贝表格行
     * @param row
     * @param newRow
     */
    public static void copyRow(XWPFTableRow row, XWPFTableRow newRow){
        row.getCtRow().setTrPr(newRow.getCtRow().getTrPr());
        List<XWPFTableCell> cells = row.getTableCells();
        for (int i = 0; i <cells.size() ; i++) {
            XWPFTableCell cell = cells.get(i);
            if(newRow.getTableCells().size()>i){
                XWPFTableCell newCell = newRow.getTableCells().get(i);
                newCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());
                copyParagraph(newCell, cell.getParagraphs());
            }else {
                XWPFTableCell newCell = newRow.addNewTableCell();
                newCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());
                copyParagraph(newCell, cell.getParagraphs());
            }

        }
    }

    /**
     * 拷贝表格列
     * @param cell
     * @param paragraphs
     */
    private static void copyParagraph(XWPFTableCell cell, List<XWPFParagraph> paragraphs) {
        for (int j = 0; j < paragraphs.size(); j++) {
            XWPFParagraph paragraph = paragraphs.get(j);
            if(cell.getParagraphs().size() > j){
                XWPFParagraph newParagraph = cell.getParagraphs().get(j);
                newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
                copyRun(newParagraph, paragraph.getRuns());
            }else {
                XWPFParagraph newParagraph = cell.addParagraph();
                newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
                copyRun(newParagraph, paragraph.getRuns());
            }
        }
    }

    /**
     * 拷贝RUN
     * @param paragraph
     * @param runs
     */
    public static void copyRun(XWPFParagraph paragraph, List<XWPFRun> runs) {
        for (int k = 0; k <runs.size() ; k++) {
            XWPFRun run = runs.get(k);
            if(paragraph.getRuns().size()>k){
                XWPFRun newRun = paragraph.getRuns().get(k);
                newRun.setText(run.toString());
                newRun.getCTR().setRPr(run.getCTR().getRPr());
            }else {
                XWPFRun newRun =  paragraph.insertNewRun(k);
                newRun.setText(run.toString());
                newRun.getCTR().setRPr(run.getCTR().getRPr());
            }
        }
    }

    /**
     * 插入表格
     * @param table
     * @param dataList
     * @param margeTables
     */
    public static void insertTable(XWPFTable table, List<String[]> dataList, List<TableCustomParagraphBuild.MargeTable> margeTables) {

        List<XWPFTableRow> rows =table.getRows();
        for (int i=0; i< dataList.size(); i++){
            String[] datas = dataList.get(i);
            XWPFTableRow row = null;
            if(rows.size() > i){
                row = rows.get(i);
            }else{
                row = table.insertNewTableRow(i);
            }
            List<XWPFTableCell> cells = row.getTableCells();
            for(int y=0; y< datas.length; y++){
                String data = datas[y];
                XWPFTableCell cell = null;
                if(cells.size() > y){
                    cell = cells.get(y);
                }else {
                    cell = row.createCell();
                }
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                if (paragraphs.size() > 0){
                    XWPFParagraph paragraph = paragraphs.get(0);
                    replaceParagraph(paragraph, data, true);
                }else {
                    XWPFParagraph paragraph = cell.addParagraph();
                    replaceParagraph(paragraph, data, true);
                }

            }
        }
        mergeTable(table, margeTables);
    }


    /**
     * 段落设置字符，不使用变量替换
     * @param paragraph
     * @param data
     */
    public static void replaceParagraph(XWPFParagraph paragraph, String data, boolean isTable) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = runs.size()-1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        insertText(paragraph, data, 0, isTable);
    }

    /**
     * 段落设置字符，使用变量替换
     * @param paramDate
     * @param paragraph
     */
    public static void replaceParagraph(Map<String, String> paramDate, XWPFParagraph paragraph, boolean isTable) {
        //判断此段落时候需要进行替换
        while(checkText(paragraph.getText())){
            int s = -1;
            int e = -1;
            List<String> tmpArr = new ArrayList<>();
            List<XWPFRun> runs = paragraph.getRuns();
            if(runs.size() == 1){
                tmpArr.add(runs.get(0).toString());
                s = 0;
                e = 0;
            }else {
                for (int i=0;i<runs.size();i++){
                    XWPFRun run = runs.get(i);
                    String tmpText = run.toString();
                    if (tmpText.contains("$")){
                        if (s == -1){
                            s = i;
                        }else if(s > -1){
                            e = i;
                        }
                    }
                    if (s > -1){
                        tmpArr.add(tmpText);
                    }
                    if(e > -1){
                        break;
                    }
                }
                if (e<s){
                    e = s;
                }
            }
            String value = StringUtils.join(tmpArr, "");
            if (StringUtils.isNotBlank(value)){
                String text = changeValue(value, paramDate);
                insertText(paragraph, text, e+1, isTable);
            }
            for(int i = e; i>=s; i --){
                paragraph.removeRun(i);
            }
        }
    }



    /**
     * 表格合并
     * @param table
     */
    public static void mergeTable(XWPFTable table, List<TableCustomParagraphBuild.MargeTable> margeTables){
        if (margeTables == null || margeTables.isEmpty()){
            return;
        }
        for (int i=0;i<margeTables.size();i++){
            TableCustomParagraphBuild.MargeTable margeTable = margeTables.get(i);
            TableCustomParagraphBuild.MargeTableIndex sMargeTableIndex = margeTable.getS();
            TableCustomParagraphBuild.MargeTableIndex eMargeTableIndex = margeTable.getE();
            for(int r = sMargeTableIndex.getRows();r <= eMargeTableIndex.getRows(); r++){
                mergeCellsHorizontal(table, r, sMargeTableIndex.getColumns(), eMargeTableIndex.getColumns());
            }
            mergeCellsVertically(table, sMargeTableIndex.getColumns(), sMargeTableIndex.getRows(), eMargeTableIndex.getRows());
        }
    }

    /**
     * 表格行合并
     * @param table
     * @param col
     * @param fromRow
     * @param toRow
     */
    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        if (fromRow == toRow){
            return;
        }
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if ( rowIndex == fromRow ) {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }


    /**
     * 表格列合并
     * @param table
     * @param row
     * @param fromCell
     * @param toCell
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        if (fromCell == toCell){
            return;
        }
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if ( cellIndex == fromCell ) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

}
