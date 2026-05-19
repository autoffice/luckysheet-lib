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

/**
 * 冻结类型枚举, 定义工作表冻结行列的方式.
 *
 * <p>对应 Luckysheet frozen 配置中的 type 字段, 支持首行/首列/行列/选区冻结等模式.</p>
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#frozen">Luckysheet 冻结文档</a>
 */
@AllArgsConstructor
public enum FrozenType {
    /**
     * 冻结首行
     */
    ROW("row"),

    /**
     * 冻结首列
     */
    COLUMN("column"),

    /**
     * 冻结行列
     */
    BOTH("both"),

    /**
     * 冻结行到选区
     */
    RANGE_ROW("rangeRow"),

    /**
     * 冻结列到选区
     */
    RANGE_COLUMN("rangeColumn"),

    /**
     * 冻结行列到选区
     */
    RANGE_BOTH("rangeBoth"),

    /**
     * 取消冻结
     */
    CANCEL("cancel");


    @JsonValue
    private final String lsValue;
}
