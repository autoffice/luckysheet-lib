package io.github.autoffice.luckysheet.model.sheet;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

/**
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo"></a>
 */
@AllArgsConstructor
public enum BorderRangeType {
    /**
     * 区域
     */
    RANGE("range"),
    /**
     * 单个单元格
     */
    CELL("cell");

    @JsonValue
    private final String lsValue;
}
