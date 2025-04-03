package com.helloaldis.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

/**
 * 单元格值Type类型
 *
 * @see <a
 * href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%8D%95%E5%85%83%E6%A0%BC%E5%80%BC%E6%A0%BC%E5%BC%8F">
 * cellFormatType </a>
 */
@AllArgsConstructor
public enum CellTypeEnum {
    /**
     * 自动格式
     */
    GENERAL("g"),
    /**
     * 纯文本字符串
     */
    STRING("s"),
    /**
     * 数字或货币
     */
    NUMBER("n"),
    /**
     * 日期时间
     */
    DATETIME("d"),

    /**
     * 富文本
     */
    INLINESTR("inlineStr"),

    /**
     * 未知类型
     */
    B("b"),

    /**
     * 未知类型
     */
    E("e");

    @JsonValue
    private final String lsValue;
}
