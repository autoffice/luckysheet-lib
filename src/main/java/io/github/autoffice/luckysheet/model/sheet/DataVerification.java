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

/**
 * 单元格数据验证 (Data Verification) 规则.
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#dataverification">
 *      luckysheet dataVerification</a>
 */
@Data
public class DataVerification {
    /**
     * 校验类型: dropdown, checkbox, number, text_content, text_length, date, validity
     */
    private String type;

    /**
     * 子类型: bw / nbw / eq / ne / gt / lt / gte / lte / include / exclude / card / phone 等.
     */
    private String type2;

    /**
     * 主参数, 如下拉列表逗号分隔值, 数字范围起始值, 日期起始值等.
     */
    private String value1;

    /**
     * 次参数, 如范围校验的结束值.
     */
    private String value2;

    /**
     * 是否已勾选 (checkbox 类型专用).
     */
    private Boolean checked;

    /**
     * 是否远程数据源.
     */
    private Boolean remote;

    /**
     * 是否禁止录入非法值 (true 等同 Excel 的 stop).
     */
    private Boolean prohibitInput;

    /**
     * 是否显示提示.
     */
    private Boolean hintShow;

    /**
     * 提示文字.
     */
    private String hintText;
}
