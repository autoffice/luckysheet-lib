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

/**
 * 缩小字体填充 (Shrink to Fit)
 * 当单元格内容超出单元格宽度时，自动缩小字体以适应单元格
 *
 * @author hanbd
 */
@AllArgsConstructor
@Getter
public enum ShrinkToFit {
    /**
     * 不缩小字体
     */
    NO(0),
    /**
     * 缩小字体填充
     */
    YES(1);

    @JsonValue
    private final Integer lsValue;

    /**
     * 从POI的boolean值转换
     *
     * @param shrinkToFit POI的shrinkToFit值
     * @return ShrinkToFit枚举
     */
    public static ShrinkToFit of(boolean shrinkToFit) {
        return shrinkToFit ? YES : NO;
    }

    /**
     * 转换为POI的boolean值
     *
     * @return POI使用的boolean值
     */
    public boolean isPoiValue() {
        return this == YES;
    }
}
