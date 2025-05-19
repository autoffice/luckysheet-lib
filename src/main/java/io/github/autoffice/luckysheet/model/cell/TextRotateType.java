package io.github.autoffice.luckysheet.model.cell;

import com.fasterxml.jackson.annotation.JsonValue;
import io.github.autoffice.luckysheet.util.Util;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 字体旋转类型,<b>参考微软excel中 开始->对齐方式->方向</b>
 */
@AllArgsConstructor
public enum TextRotateType {
    /**
     * 不旋转
     */
    NONE(0, (short) 0),
    /**
     * 顺时针45度
     */
    CLOCKWISE_45(1, (short) 45),
    /**
     * 逆时针45度
     */
    ANTICLOCKWISE_45(2, (short) 135),
    /**
     * 纵向。MS-excel 竖排文字
     */
    VERTICAL(3, (short) 90),
    /**
     * 顺时针90度。MS-excel 向下旋转文字
     */
    CLOCKWISE_90(4, (short) 90),
    /**
     * 逆时针90度。 MS-excel 向上旋转文字
     */
    ANTICLOCKWISE_90(5, (short) 180);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final short poiValue;

    private static final Map<Short, TextRotateType> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(TextRotateType::getPoiValue, Function.identity(),
                    (existing, replacement) -> existing));

    public static TextRotateType of(short rotation) {
        TextRotateType textRotateType = TYPES.get(rotation);
        return Util.requireNonNullElse(textRotateType, NONE);
    }
}
