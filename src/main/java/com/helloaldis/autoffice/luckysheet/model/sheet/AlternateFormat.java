package com.helloaldis.autoffice.luckysheet.model.sheet;

import lombok.Data;

@Data
public class AlternateFormat {
    /**
     * 单元格范围
     */
    private Range cellrange;
    /**
     *
     */
    private AlternateFormatValue format;
    /**
     * 含有页眉
     */
    private Boolean hasRowHeader;
    /**
     * 含有页脚
     */
    private Boolean hasRowFooter;

    @Data
    public static class AlternateFormatValue {
        /**
         * 页眉颜色
         */
        private AlternateFormatColor head;
        /**
         * 第一种颜色
         */
        private AlternateFormatColor one;
        /**
         * 第二种颜色
         */
        private AlternateFormatColor two;
        /**
         * 页脚颜色
         */
        private AlternateFormatColor foot;
    }

    @Data
    public static class AlternateFormatColor {
        private String fc;
        private String bc;
    }
}
