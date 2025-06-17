package io.github.autoffice.luckysheet.model.sheet;

import lombok.Data;

import java.util.List;

/**
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#calcchain">luckysheet官方文档</a>
 */
@Data
public class CalcChain {
    /**
     * 行数
     */
    private Integer r;
    /**
     * 列数
     */
    private Integer c;
    /**
     * 工作表id
     */
    private Integer index;
    /**
     * 公式信息，包含公式计算结果和公式字符串
     */
    private List<Object> func;
    /**
     * "w"：采用深度优先算法 "b":普通计算
     */
    private String color;
}
