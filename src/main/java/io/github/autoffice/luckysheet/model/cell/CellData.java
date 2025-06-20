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

import lombok.Data;

/**
 * @see <a
 * href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC">
 * 单元格</a>
 */
@Data
public class CellData {
    /**
     * 行index
     */
    private Integer r;
    /**
     * 列index
     */
    private Integer c;
    /**
     * 单元格值
     */
    private CellValue v;
}
