package com.helloaldis.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@AllArgsConstructor
public enum Cancelline {
    /**
     * 常规
     */
    NORMAL(0, false),
    /**
     * 加粗
     */
    CANCELLINE(1, true);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final boolean poiValue;

    private static final Map<Boolean, Cancelline> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(Cancelline::isPoiValue, Function.identity()));

    public static Cancelline of(boolean strikeout) {
        return TYPES.get(strikeout);
    }
}
