package com.helloaldis.autoffice.luckysheet.model.cell;

import lombok.Data;

/**
 * @see <a
 * href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC">
 * 单元格</a>
 */
@Data
public class CellData {
    /**
     * 行index
     */
    private Integer r;
    /**
     * 列index
     */
    private Integer c;
    /**
     * 单元格值
     */
    private CellValue v;
}
