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

/**
 * 文本换行类型
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum TextBreakType {
    /**
     * 文本截断
     */
    TRUNCATION(0),
    /**
     * 文本溢出
     */
    OVERFLOW(1),
    /**
     * 自动换行
     */
    LINE_WRAP(2);

    @JsonValue
    private final Integer lsValue;
}
