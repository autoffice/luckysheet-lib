package com.helloaldis.autoffice.luckysheet.model.image;

import lombok.Data;

@Data
public class ImageBorder {

    /**
     * 边框宽度
     */
    private Integer width;

    /**
     * 边框样式类型
     */
    private ImageBorderType style;

    /**
     * 边框验收
     */
    private String color;
}
