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

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Luckysheet 支持的字体类型
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC">
 * 基本单元格 ff</a>
 */
@AllArgsConstructor
public enum FontFamily {
    /**
     * Times New Roman
     */
    TIMES_NEW_ROMAN(0, "Times New Roman"),
    /**
     * Arial
     */
    ARIAL(1, "Arial"),
    /**
     * Tahoma
     */
    TAHOMA(2, "Tahoma"),
    /**
     * Verdana
     */
    VERDANA(3, "Verdana"),
    /**
     * 微软雅黑
     */
    MS_YA_HEI(4, "微软雅黑"),
    /**
     * 宋体
     */
    SONG(5, "宋体"),
    /**
     * 黑体
     */
    ST_HEI_TI(6, "黑体"),
    /**
     * 楷体
     */
    ST_KAI_TI(7, "楷体"),
    /**
     * 仿宋
     */
    ST_FANG_SONG(8, "仿宋"),
    /**
     * 新宋体
     */
    ST_SONG(9, "新宋体"),
    /**
     * 华文新魏
     */
    HUA_WEN_XIN_WEI(10, "华文新魏"),
    /**
     * 华文行楷
     */
    HUA_WEN_XING_KAI(11, "华文行楷"),
    /**
     * 华文隶书
     */
    HUA_WEN_LI_SHU(12, "华文隶书");

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final String poiValue;

    private static final Map<String, FontFamily> NAMES = Arrays.stream(values())
            .collect(Collectors.toMap(FontFamily::getPoiValue, Function.identity()));

    /**
     * @param name 字体名
     * @return 如果无对应字体, 默认返回 {@link FontFamily#ARIAL}
     */
    @JsonCreator
    public static FontFamily of(String name) {
        FontFamily fontFamily = NAMES.get(name);
        if (Objects.isNull(fontFamily)) {
            return ARIAL;
        }
        return fontFamily;
    }
}
