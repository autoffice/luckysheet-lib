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
