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
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#calcchain">luckysheet官方文档</a>
 */
@Data
public class CalcChain {
    /**
     * 行数
     */
    private Integer r;
    /**
     * 列数
     */
    private Integer c;
    /**
     * 工作表id
     */
    private Integer index;
    /**
     * 公式信息，包含公式计算结果和公式字符串
     */
    private List<Object> func;
    /**
     * "w"：采用深度优先算法 "b":普通计算
     */
    private String color;
}
