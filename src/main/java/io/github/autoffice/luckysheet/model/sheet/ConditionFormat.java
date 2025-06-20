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


/**
 * luckysheet 条件格式配置信息
 *
 * @see <a
 * href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#luckysheet-conditionformat-save">
 * luckysheet-conditionformat-save</a>
 */
@Data
public class ConditionFormat {
    /**
     * 突出显示单元格规则和项目选区规则
     */
    private ConditionFormatType type;
    /**
     * 条件应用范围
     */
    private List<Range> cellrange;
    /**
     * 格式,根据不同的{@link ConditionFormat#type},实际对象类型不同
     */
    private Object format;
    /**
     * 类型
     *
     * <p>Detailed settings,comparison parameters
     */
    private String conditionName;
    /**
     * 条件值所在单元格
     *
     * <p>Detailed settings,comparison range
     */
    private List<Range> conditionRange;
    /**
     * 自定义传入的条件值
     */
    private List<Object> conditionValue;
}
