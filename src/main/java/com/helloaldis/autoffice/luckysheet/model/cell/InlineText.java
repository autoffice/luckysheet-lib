package com.helloaldis.autoffice.luckysheet.model.cell;

import lombok.Data;

@Data
public class InlineText {
    /**
     * 字体
     */
    private FontFamily ff;
    /**
     * 字体颜色. #fff000
     */
    private String fc;
    /**
     * 字体大小
     */
    private Short fs;

    /**
     * 删除线. 0 常规 、 1 删除线
     */
    private Cancelline cl;
    /**
     * 下划线. 0 无 、 1 有
     */
    private Underline un;
    /**
     * 粗体. 0 常规 、 1加粗
     */
    private Bold bl;
    /**
     * 斜体. 0 常规 、 1 斜体
     */
    private Italic it;

    /**
     * 文本内容
     */
    private String v;
}
