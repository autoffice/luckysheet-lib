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

import io.github.autoffice.luckysheet.model.cell.MergeCell;
import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 表格行高、列宽、合并单元格、边框、隐藏行等设置
 * <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#config"></a>
 */
@Data
public class SheetConfig {
    /**
     * new 对象,并对各字段初始化赋空值
     */
    public SheetConfig() {
        this.merge = new HashMap<>();
        this.borderInfo = new ArrayList<>();
        this.rowlen = new HashMap<>();
        this.columnlen = new HashMap<>();
        this.colhidden = new HashMap<>();
        this.rowhidden = new HashMap<>();
        this.customHeight = new HashMap<>();
        this.customWidth = new HashMap<>();
    }

    /**
     * 合并单元格设置
     * 对象中的key为r + '_' + c的拼接值，value为左上角单元格信息: r:行数，c:列数，rs：合并的行数，cs:合并的列数
     */
    private Map<String, MergeCell> merge;

    /**
     * 单元格的边框信息
     */
    private List<Border> borderInfo;

    /**
     * 每行单元格的行高. key: row index, value: row height
     */
    private Map<Integer, Integer> rowlen;

    /**
     * 每行单元格的列宽信息. key: column index, value: column width
     */
    private Map<Integer, Integer> columnlen;

    /**
     * 隐藏行信息，格式为：rowhidden[行数]: 0. key: index of hidden row, value: 0,
     * key指定行数即可，value总是为0
     */
    private Map<Integer, Integer> rowhidden;

    /**
     * 隐藏列 格式为：colhidden[列数]: 0
     */
    private Map<Integer, Integer> colhidden;

    private Map<Integer, Integer> customHeight;

    private Map<Integer, Integer> customWidth;

    /**
     * 工作表保护，可以设置当前整个工作表不允许编辑或者部分区域不可编辑，
     * 如果要申请编辑权限需要输入密码，自定义配置用户可以操作的类型等。
     */
    private Object authority;
}
