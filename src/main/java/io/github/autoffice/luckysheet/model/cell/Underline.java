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
