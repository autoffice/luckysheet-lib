package com.helloaldis.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@AllArgsConstructor
public enum Bold {
    /**
     * 常规
     */
    NORMAL(0, false),
    /**
     * 加粗
     */
    BOLD(1, true);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final boolean poiValue;

    private static final Map<Boolean, Bold> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(Bold::isPoiValue, Function.identity()));

    public static Bold of(boolean poiValue) {
        return TYPES.get(poiValue);
    }
}
