package com.helloaldis.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

/**
 * 文本换行类型
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum TextBreakType {
    /**
     * 文本截断
     */
    TRUNCATION(0),
    /**
     * 文本溢出
     */
    OVERFLOW(1),
    /**
     * 自动换行
     */
    LINE_WRAP(2);

    @JsonValue
    private final Integer lsValue;
}
