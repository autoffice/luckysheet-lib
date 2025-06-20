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
