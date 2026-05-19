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
 * 交替颜色配置, 用于设置工作表中行的交替背景色.
 *
 * <p>对应 Luckysheet JSON 中的 luckysheet_alternateformat_save 数组元素.</p>
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#luckysheet-alternateformat-save">Luckysheet 交替颜色文档</a>
 */
@Data
public class AlternateFormat {
    /**
     * 单元格范围
     */
    private Range cellrange;
    /**
     *
     */
    private AlternateFormatValue format;
    /**
     * 含有页眉
     */
    private Boolean hasRowHeader;
    /**
     * 含有页脚
     */
    private Boolean hasRowFooter;

    /**
     * 交替颜色值, 包含页眉、页脚和两种交替行的颜色配置.
     */
    @Data
    public static class AlternateFormatValue {
        /**
         * 页眉颜色
         */
        private AlternateFormatColor head;
        /**
         * 第一种颜色
         */
        private AlternateFormatColor one;
        /**
         * 第二种颜色
         */
        private AlternateFormatColor two;
        /**
         * 页脚颜色
         */
        private AlternateFormatColor foot;
    }

    /**
     * 交替颜色, 包含字体颜色 (fc) 和背景颜色 (bc).
     */
    @Data
    public static class AlternateFormatColor {
        private String fc;
        private String bc;
    }
}
