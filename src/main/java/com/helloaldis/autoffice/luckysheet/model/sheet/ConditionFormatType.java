package com.helloaldis.autoffice.luckysheet.model.sheet;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public enum ConditionFormatType {
    /**
     * 突出显示单元格规则和项目选区规则
     */
    DEFAULT("default"),
    /**
     * 单个单元格
     */
    DATA_BAR("dataBar"),
    /**
     * 图标集
     */
    ICONS("icons"),
    /**
     * 色阶
     */
    COLOR_GRADATION("colorGradation");

    @JsonValue
    private final String lsValue;
}
