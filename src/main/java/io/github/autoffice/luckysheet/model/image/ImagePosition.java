package io.github.autoffice.luckysheet.model.image;

import lombok.Data;

@Data
public class ImagePosition {
    /**
     * 图片展示宽度
     */
    private Integer width;
    /**
     * 图片展示高度
     */
    private Integer height;
    /**
     * 图片离表格左边的位置
     */
    private Integer left;
    /**
     * 图片离表格顶部的位置
     */
    private Integer top;
}
