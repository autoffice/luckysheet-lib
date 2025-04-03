package com.helloaldis.autoffice.luckysheet.model.sheet;

import lombok.Data;

import java.util.List;

/**
 * 单元格边框. {@link BorderRangeType#CELL}
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%90%AB%E8%BE%B9%E6%A1%86%E5%8D%95%E5%85%83%E6%A0%BC" />
 */
@Data
public class Border {
    /**
     * 边框范围类型
     */
    private BorderRangeType rangeType;

    private Value value;

    private BorderType borderType;

    private BorderStyleType style;

    private String color;

    private List<Range> range;

    @Data
    public static class Value {
        /**
         * 单元格行索引值
         */
        private Integer row_index;
        /**
         * 单元格列索引值
         */
        private Integer col_index;
        /**
         * 左侧边框样式
         */
        private Style l;
        /**
         * 右侧边框样式
         */
        private Style r;
        /**
         * 顶部边框样式
         */
        private Style t;
        /**
         * 底部边框样式
         */
        private Style b;
    }

    @Data
    public static class Style {
        /**
         * 边框线格式
         */
        private BorderStyleType style;
        /**
         * 颜色. 格式: {@code rgb(255,0,0)}
         */
        private String color;
    }
}

