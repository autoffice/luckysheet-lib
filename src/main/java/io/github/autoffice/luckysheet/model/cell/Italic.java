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
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@AllArgsConstructor
public enum Italic {
    /**
     * 常规
     */
    NORMAL(0, false),
    /**
     * 加粗
     */
    ITALIC(1, true);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final boolean poiValue;

    private static final Map<Boolean, Italic> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(Italic::isPoiValue, Function.identity()));

    public static Italic of(boolean italic) {
        return TYPES.get(italic);
    }
}
