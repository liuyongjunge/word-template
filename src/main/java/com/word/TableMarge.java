package com.word;

import lombok.Getter;

import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/26 11:31
 */
public class TableMarge {

    /**
     *
     * 注入合并行列数据
     * @return
     */
    public Map<String, List<MargeTable>> getMargeData(){
        return Collections.emptyMap();
    }


    /**
     * 表格合并范围
     * @param rows 合并行
     * @param columns 合并列
     * @return
     */
    public static MargeTableIndex instanceMargeTableIndex(int rows, int columns){
        return new MargeTableIndex(rows, columns);
    }

    /**
     * 表格合并数据
     * @param sMargeTableIndex  合并开始范围
     * @param eMargeTableIndex  合并结束范围
     * @return
     */
    public static MargeTable instanceMargeTable(MargeTableIndex sMargeTableIndex, MargeTableIndex eMargeTableIndex){
        return new MargeTable(sMargeTableIndex, eMargeTableIndex);
    }

    /**
     * 表格合并数据
     * @param sRows 合并开始行
     * @param eRows 合并结束行
     * @param sColumns 合并开始列
     * @param eColumns 合并结束列
     * @return
     */
    public static MargeTable instanceMargeTable(int sRows, int eRows, int sColumns, int eColumns){
        return new MargeTable(instanceMargeTableIndex(sRows, sColumns), instanceMargeTableIndex(eRows, eColumns));
    }


    @Getter
    static class MargeTable{

        /**
         * 合并行列的起始范围
         */
        private MargeTableIndex s;

        /**
         * 合并行列的结束范围
         */
        private MargeTableIndex e;

        private MargeTable(MargeTableIndex s, MargeTableIndex e){
            this.s = s;
            this.e = e;
        }
    }

    @Getter
    static class MargeTableIndex{

        /**
         * 合并行的范围
         */
        private int rows;

        /**
         * 合并列的范围
         */
        private int columns;

        private MargeTableIndex(int rows, int columns){
            this.rows = rows;
            this.columns =columns;
        }
    }
}
