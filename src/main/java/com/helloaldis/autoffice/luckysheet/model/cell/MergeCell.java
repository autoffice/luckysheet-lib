package com.helloaldis.autoffice.luckysheet.model.cell;

import lombok.Data;

/**
 * luckysheet中合并单元格的表示
 * <p>
 * +-----------------+       +----------------+
 * |(0,0)   |(0,1)   |       |                |
 * +-----------------+ +---> |                |
 * |(1,0)   |(1,1)   |       |                |
 * +-----------------+       +----------------+
 * <p>
 * 如上4个单元格合并后: startRow=0,startCol=0,rowsNum=2,colsNum=2
 * <p>
 * href="<a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC">...</a>">
 * luckysheet合并单元格 </a>
 */
@Data
public class MergeCell {
    /**
     * 合并单元格主单元格行号, <b>0 based</b>
     */
    private Integer r;
    /**
     * 合并单元格主单元格列号, <b>0 based</b>
     */
    private Integer c;
    /**
     * 合并单元格所占行数
     */
    private Integer rs;
    /**
     * 合并单元格所占列数
     */
    private Integer cs;
}
