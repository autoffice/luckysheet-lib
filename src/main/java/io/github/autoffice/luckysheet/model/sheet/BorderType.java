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
 * 边框类型枚举, 定义边框应用的范围和方向.
 *
 * <p>对应 Luckysheet config.borderInfo 中 borderType 字段的取值.</p>
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">Luckysheet 边框文档</a>
 */
@AllArgsConstructor
public enum BorderType {
    /**
     *
     */
    LEFT("border-left"),
    /**
     *
     */
    RIGHT("border-right"),
    /**
     *
     */
    TOP("border-top"),
    /**
     *
     */
    BOTTOM("border-bottom"),
    /**
     *
     */
    ALL("border-all"),
    /**
     *
     */
    OUTSIDE("border-outside"),
    /**
     *
     */
    INSIDE("border-inside"),
    /**
     *
     */
    HORIZONTAL("border-horizontal"),
    /**
     *
     */
    VERTICAL("border-vertical"),
    /**
     *
     */
    NONE("border-none");

    @JsonValue
    private final String lsValue;
}
