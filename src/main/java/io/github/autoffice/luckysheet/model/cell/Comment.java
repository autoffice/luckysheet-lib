package io.github.autoffice.luckysheet.model.cell;

import lombok.Data;

/**
 * 批注
 * <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC"></a>
 *
 * @author hanbd
 */
@Data
public class Comment {
    /**
     * 批注框左边距
     */
    private Integer left;
    /**
     * 批注框上边距
     */
    private Integer top;
    /**
     * 批注框宽度
     */
    private Integer width;
    /**
     * 批注框高度
     */
    private Integer height;
    /**
     * 批注内容
     */
    private String value;
    /**
     * 批注框是否显示
     */
    private Boolean isshow;
}
