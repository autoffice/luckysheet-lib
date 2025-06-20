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

    @Data
    public static class AlternateFormatColor {
        private String fc;
        private String bc;
    }
}
