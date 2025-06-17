package io.github.autoffice.luckysheet.model.sheet;

import lombok.Data;

import java.util.List;

/**
 * 选中的区域，支持多选，是一个包含多个选区对象的一维数组
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#luckysheet-select-save">luckysheet官方文档</a>
 */
@Data
public class Range {

    /**
     * 行起止index范围
     */
    private List<Integer> row;
    /**
     * 列起止index范围
     */
    private List<Integer> column;


}
