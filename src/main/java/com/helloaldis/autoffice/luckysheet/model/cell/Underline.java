package com.helloaldis.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.FontUnderline;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@AllArgsConstructor
public enum Underline {
    /**
     * 常规
     */
    NORMAL(0, FontUnderline.NONE),
    /**
     * 加粗
     */
    UNDERLINE(1, FontUnderline.SINGLE);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final FontUnderline poiValue;

    private static final Map<FontUnderline, Underline> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(Underline::getPoiValue, Function.identity()));

    public static Underline of(FontUnderline fontUnderline) {
        Underline underline = TYPES.get(fontUnderline);
        if (underline == null) {
            return NORMAL;
        }

        return underline;
    }
}
