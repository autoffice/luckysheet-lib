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
 * 单元格值Type类型
 *
 * @see <a
 * href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%8D%95%E5%85%83%E6%A0%BC%E5%80%BC%E6%A0%BC%E5%BC%8F">
 * cellFormatType </a>
 */
@AllArgsConstructor
public enum CellTypeEnum {
    /**
     * 自动格式
     */
    GENERAL("g"),
    /**
     * 纯文本字符串
     */
    STRING("s"),
    /**
     * 数字或货币
     */
    NUMBER("n"),
    /**
     * 日期时间
     */
    DATETIME("d"),

    /**
     * 富文本
     */
    INLINESTR("inlineStr"),

    /**
     * 未知类型
     */
    B("b"),

    /**
     * 未知类型
     */
    E("e");

    @JsonValue
    private final String lsValue;
}
