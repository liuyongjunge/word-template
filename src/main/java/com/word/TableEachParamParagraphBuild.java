package com.word;

import com.word.ParagraphBuild;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.math.BigInteger;
import java.util.List;
import java.util.Map;
import java.util.Timer;
import java.util.TimerTask;

/**
 * 存在表格模型、在表格中存在 ##{**}##标记
 * 循环替换表格中的变量
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/18 11:01
 */
public abstract class TableEachParamParagraphBuild extends TableMarge implements ParagraphBuild {

    @Override
    public void replaceParagraph(XWPFDocument document) {

        Map<String,List<Map<String, String>>> paramData = getParamData();
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {

            XWPFTable table = tables.get(i);
            String tableKey = ParagraphBuildHandle.checkTableEach(table);
            if(StringUtils.isNotBlank(tableKey) && paramData.containsKey(tableKey)
                    && ParagraphBuildHandle.checkText(table.getText())){
                List<XWPFTableRow> rows = table.getRows();
                Map<String, List<MargeTable>> margeData = getMargeData();
                ParagraphBuildHandle.eachTableData(paramData.get(tableKey), table, margeData.get(tableKey));
            }
        }
    }

    public void removeRow(XWPFTable table, int pos){
        new Timer().schedule(new TimerTask() {
            @Override
            public void run() {
                table.removeRow(pos);
            }
        },100L);
    }


    /**
     *
     * 注入业务数据
     * @return
     */
    public abstract Map<String,List<Map<String, String>>> getParamData();
}
