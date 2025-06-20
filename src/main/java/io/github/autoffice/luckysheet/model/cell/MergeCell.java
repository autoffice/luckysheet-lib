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
 * luckysheet中合并单元格的表示
 * <p>
 * +-----------------+       +----------------+
 * |(0,0)   |(0,1)   |       |                |
 * +-----------------+ +--- |                |
 * |(1,0)   |(1,1)   |       |                |
 * +-----------------+       +----------------+
 * <p>
 * luckysheet合并单元格，如上4个单元格合并后: startRow=0,startCol=0,rowsNum=2,colsNum=2
 * <p>
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC">luckysheet官方文档</a>
 */
@Data
public class MergeCell {
    /**
     * 合并单元格主单元格行号, <b>0 based</b>
     */
    private Integer r;
    /**
     * 合并单元格主单元格列号, <b>0 based</b>
     */
    private Integer c;
    /**
     * 合并单元格所占行数
     */
    private Integer rs;
    /**
     * 合并单元格所占列数
     */
    private Integer cs;
}
