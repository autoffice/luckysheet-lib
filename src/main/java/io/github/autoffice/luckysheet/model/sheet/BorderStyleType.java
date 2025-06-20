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
package io.github.autoffice.luckysheet.model.sheet;

import com.fasterxml.jackson.annotation.JsonValue;
import io.github.autoffice.luckysheet.util.Util;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.BorderStyle;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 边框线格式
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">config-borderInfo</a>
 */
@AllArgsConstructor
public enum BorderStyleType {
    /**
     * 无
     */
    NONE(0, "None", BorderStyle.NONE),
    /**
     * Thin
     */
    THIN(1, "Thin", BorderStyle.THIN),
    /**
     * Hair
     */
    HAIR(2, "Hair", BorderStyle.HAIR),
    /**
     * Dotted
     */
    DOTTED(3, "Dotted", BorderStyle.DOTTED),
    /**
     * Dashed
     */
    DASHED(4, "Dashed", BorderStyle.DASHED),
    /**
     * DashDot
     */
    DASH_DOT(5, "DashDot", BorderStyle.DASH_DOT),
    /**
     * DashDotDot
     */
    DASH_DOT_DOT(6, "DashDotDot", BorderStyle.DASH_DOT_DOT),
    /**
     * Double
     */
    DOUBLE(7, "Double", BorderStyle.DOUBLE),
    /**
     * Medium
     */
    MEDIUM(8, "Medium", BorderStyle.MEDIUM),
    /**
     * MediumDashed
     */
    MEDIUM_DASHED(9, "MediumDashed", BorderStyle.MEDIUM_DASHED),
    /**
     * MediumDashDot
     */
    MEDIUM_DASH_DOT(10, "MediumDashDot", BorderStyle.MEDIUM_DASH_DOT),
    /**
     * MediumDashDotDot
     */
    MEDIUM_DASH_DOT_DOT(11, "MediumDashDotDot", BorderStyle.MEDIUM_DASH_DOT_DOT),
    /**
     * SlantedDashDot
     */
    SLANTED_DASH_DOT(12, "SlantedDashDot", BorderStyle.SLANTED_DASH_DOT),
    /**
     * Thick
     */
    THICK(13, "Thick", BorderStyle.THICK);

    /**
     * 格式index
     */
    @JsonValue
    private final Integer lsValue;
    /**
     * 格式名
     */
    @Getter
    private final String name;
    /**
     * 对应的poi border style
     */
    @Getter
    private final BorderStyle poiValue;

    private static final Map<BorderStyle, BorderStyleType> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(BorderStyleType::getPoiValue, Function.identity()));

    public static BorderStyleType of(BorderStyle borderStyle) {
        BorderStyleType borderStyleType = TYPES.get(borderStyle);
        return Util.requireNonNullElse(borderStyleType, THIN);
    }
}
