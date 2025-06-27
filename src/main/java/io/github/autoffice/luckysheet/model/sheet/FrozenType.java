package io.github.autoffice.luckysheet.model.sheet;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public enum FrozenType {
    /**
     * 冻结首行
     */
    ROW("row"),

    /**
     * 冻结首列
     */
    COLUMN("column"),

    /**
     * 冻结行列
     */
    BOTH("both"),

    /**
     * 冻结行到选区
     */
    RANGE_ROW("rangeRow"),

    /**
     * 冻结列到选区
     */
    RANGE_COLUMN("rangeColumn"),

    /**
     * 冻结行列到选区
     */
    RANGE_BOTH("rangeBoth"),

    /**
     * 取消冻结
     */
    CANCEL("cancel");


    @JsonValue
    private final String lsValue;
}
