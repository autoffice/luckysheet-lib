package io.github.autoffice.luckysheet.model.image;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

@Data
public class SheetImage {
    /**
     * 图片定位类型
     */
    private ImageType type;

    /**
     * 图片url信息，可以是base64格式的图片
     */
    private String src;

    /**
     * 图片原始宽度
     */
    private Integer originWidth;

    /**
     * 图片原始高度
     */
    private Integer originHeight;

    /**
     * 图片位置信息
     */
    @JsonProperty("default")
    private ImagePosition position;

    /**
     * 图片裁剪信息
     */
    private ImageCrop crop;

    /**
     * 固定位置
     */
    @JsonProperty("isFixedPos")
    private boolean isFixedPos;

    /**
     * 固定位置左位移
     */
    private Integer fixedLeft;
    /**
     * /固定位置顶位移
     */
    private Integer fixedTop;

    /**
     * 问题
     */
    private ImageBorder border;
}
