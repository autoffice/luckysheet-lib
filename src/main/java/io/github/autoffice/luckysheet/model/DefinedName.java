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
package io.github.autoffice.luckysheet.model;

import lombok.Data;

/**
 * Workbook 级命名范围 (Defined Name).
 */
@Data
public class DefinedName {
    /**
     * 名称.
     */
    private String name;

    /**
     * 引用公式, eg: {@code Sheet1!$A$1:$B$5}
     */
    private String formula;

    /**
     * 所属 sheet 索引, -1 表示 workbook 级.
     */
    private Integer sheetIndex;

    /**
     * 备注.
     */
    private String comment;

    /**
     * 是否隐藏.
     */
    private Boolean hidden;
}
