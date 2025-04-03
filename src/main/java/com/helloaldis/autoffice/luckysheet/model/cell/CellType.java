package com.helloaldis.autoffice.luckysheet.model.cell;

import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.List;

/**
 * 单元格值格式
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%8D%95%E5%85%83%E6%A0%BC%E5%80%BC%E6%A0%BC%E5%BC%8F">
 * 单元格值格式</a>
 */
@Data
public class CellType {
    /**
     * Format格式的定义字符串
     *
     * @see XSSFCellStyle#getDataFormatString()
     */
    private String fa;
    /**
     * Type类型
     */
    private CellTypeEnum t;

    private List<InlineText> s;
}
