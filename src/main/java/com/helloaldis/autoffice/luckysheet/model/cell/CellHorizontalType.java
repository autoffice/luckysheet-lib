package com.helloaldis.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import com.helloaldis.autoffice.luckysheet.util.Util;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 单元格的水平对齐方式
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum CellHorizontalType {
    /**
     * 居中对齐
     */
    CENTER(0, HorizontalAlignment.CENTER),
    /**
     * 左对齐
     */
    LEFT(1, HorizontalAlignment.LEFT),
    /**
     * 右对齐
     */
    RIGHT(2, HorizontalAlignment.RIGHT);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final HorizontalAlignment poiValue;

    private static final Map<HorizontalAlignment, CellHorizontalType> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(CellHorizontalType::getPoiValue, Function.identity()));

    public static CellHorizontalType of(HorizontalAlignment alignment) {
        CellHorizontalType cellHorizontalType = TYPES.get(alignment);
        return Util.requireNonNullElse(cellHorizontalType, LEFT);
    }
}
