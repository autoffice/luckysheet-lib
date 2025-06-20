/*
 * Copyright © 2025 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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
 * 字体旋转类型,<b>参考微软excel中 开始-对齐方式-方向</b>
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
