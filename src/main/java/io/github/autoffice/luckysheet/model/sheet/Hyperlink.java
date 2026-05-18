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
 * 单元格超链接设置.
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#hyperlink">luckysheet hyperlink</a>
 */
@Data
public class Hyperlink {
    /**
     * 链接类型：external 外部链接(URL/邮件), internal 内部链接(工作表引用)
     */
    private String linkType;

    /**
     * 链接地址. 外部链接为URL/邮件, 内部链接为工作表单元格引用, eg: {@code Sheet2!A1}
     */
    private String linkAddress;

    /**
     * 鼠标悬停提示
     */
    private String linkTooltip;
}
