package com.helloaldis.autoffice.luckysheet.model.image;

import lombok.Data;

@Data
public class ImageCrop {
    /**
     * 图片裁剪后宽度
     */
    private Integer width;
    /**
     * 图片裁剪后高度
     */
    private Integer height;
    /**
     * 图片裁剪后离未裁剪时左边的位移
     */
    private Integer offsetLeft;
    /**
     * 图片裁剪后离未裁剪时顶部的位移
     */
    private Integer offsetTop;
}
