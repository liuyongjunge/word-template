package com.word;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.util.List;
import java.util.Map;

/**
 * 存在表格模型、固定行、只替换表格中的变量
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/17 17:25
 */
public abstract class TableParamParagraphBuild implements ParagraphBuild {


    /**
     * 遍历表格,并替换模板
     * @param document
     */
    @Override
    public void replaceParagraph(XWPFDocument document) {
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            XWPFTable table = tables.get(i);

            String tableKey = ParagraphBuildHandle.checkTableEach(table);
            if(StringUtils.isNotBlank(tableKey)) {
                continue;
            }
            if(ParagraphBuildHandle.checkText(table.getText())){
                List<XWPFTableRow> rows = table.getRows();
                ParagraphBuildHandle.eachTableRow(getParamData(), rows.toArray(new XWPFTableRow[rows.size()]));
            }
        }
    }

    /**
     *
     * 注入业务数据
     * @return
     */
    public abstract Map<String, String> getParamData();
}
