package com.helloaldis.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import com.helloaldis.autoffice.luckysheet.util.Util;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 单元格内值的垂直对齐方式
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum CellVerticalType {
    /**
     * 居中对齐
     */
    CENTER(0, VerticalAlignment.CENTER),
    /**
     * 上对齐
     */
    TOP(1, VerticalAlignment.TOP),
    /**
     * 下对齐
     */
    BOTTOM(2, VerticalAlignment.BOTTOM);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final VerticalAlignment poiValue;

    private static final Map<VerticalAlignment, CellVerticalType> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(CellVerticalType::getPoiValue, Function.identity()));

    public static CellVerticalType of(VerticalAlignment verticalAlignment) {
        CellVerticalType cellVerticalType = TYPES.get(verticalAlignment);
        return Util.requireNonNullElse(cellVerticalType, CENTER);
    }
}
