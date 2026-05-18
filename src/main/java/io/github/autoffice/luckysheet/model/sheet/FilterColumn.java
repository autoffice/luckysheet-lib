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

import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * 自动筛选中某一列的条件设置.
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#filter">luckysheet filter</a>
 */
@Data
public class FilterColumn {
    /**
     * 过滤类型字符串, 如 textFilter / numberFilter
     */
    private String str;

    /**
     * 枚举值列表筛选 (选中的值).
     */
    private List<String> stringsArray;

    /**
     * 条件式筛选 (key 为操作符, value 为参数), 例如 {@code {"greaterThan": 100}}.
     */
    private Map<String, Object> caljs;

    /**
     * 被此筛选隐藏的行 map (key 为行号).
     */
    private Map<String, Object> rowhidden;
}
