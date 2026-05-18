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
 * 工作表保护 (authority).
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#authority">
 *      luckysheet authority</a>
 */
@Data
public class Authority {
    /**
     * 是否启用工作表保护. 1 启用; 0 不启用.
     */
    private Integer sheet;

    /**
     * 访问提示.
     */
    private String hintText;

    /**
     * 是否允许选择被锁定单元格.
     */
    private Integer selectLockedCells;

    /**
     * 是否允许选择未锁定的单元格.
     */
    private Integer selectunLockedCells;

    /**
     * 是否允许设置单元格格式.
     */
    private Integer formatCells;

    /**
     * 是否允许设置列格式.
     */
    private Integer formatColumns;

    /**
     * 是否允许设置行格式.
     */
    private Integer formatRows;

    /**
     * 是否允许插入列.
     */
    private Integer insertColumns;

    /**
     * 是否允许插入行.
     */
    private Integer insertRows;

    /**
     * 是否允许插入超链接.
     */
    private Integer insertHyperlinks;

    /**
     * 是否允许删除列.
     */
    private Integer deleteColumns;

    /**
     * 是否允许删除行.
     */
    private Integer deleteRows;

    /**
     * 是否允许排序.
     */
    private Integer sort;

    /**
     * 是否允许筛选.
     */
    private Integer filter;

    /**
     * 是否允许使用数据透视表.
     */
    private Integer usePivotTablereports;

    /**
     * 是否允许编辑对象.
     */
    private Integer editObjects;

    /**
     * 是否允许编辑方案.
     */
    private Integer editScenarios;

    /**
     * 密码 hash 算法名 (如 SHA-512).
     */
    private String algorithmName;

    /**
     * 密码盐值 (base64).
     */
    private String saltValue;

    /**
     * 密码哈希迭代次数.
     */
    private Integer spinCount;

    /**
     * 密码哈希值 (base64).
     */
    private String hashValue;
}
