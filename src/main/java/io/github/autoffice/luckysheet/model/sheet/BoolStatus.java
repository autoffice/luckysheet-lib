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
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 布尔状态枚举, 用于表示 Luckysheet 中 0/1 形式的开关属性.
 *
 * <p>对应 Luckysheet JSON 中 hide、status、showGridLines 等字段.</p>
 */
@AllArgsConstructor
public enum BoolStatus {
    /**
     * false 0
     */
    FALSE(0, false),
    /**
     * true 1
     */
    TRUE(1, true);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final boolean poiValue;

    /**
     * 从 POI 布尔值转换.
     *
     * @param poiValue POI 中的布尔值
     * @return 对应的枚举值
     */
    public static BoolStatus of(boolean poiValue) {
        return poiValue ? TRUE : FALSE;
    }
}
