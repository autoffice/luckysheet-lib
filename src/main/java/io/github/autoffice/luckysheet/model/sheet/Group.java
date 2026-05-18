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
 * 行列分组(大纲)的单个分组区段.
 *
 * <p>Excel xlsx 中每行/列通过 outlineLevel 表达层级, 通过 collapsed 与隐藏状态
 * 表达是否折叠. 为便于 Luckysheet 侧使用, 本模型将连续同层级的行或列打包成区段.</p>
 */
@Data
public class Group {
    /**
     * 区段起始行/列索引 (含).
     */
    private Integer start;

    /**
     * 区段结束行/列索引 (含).
     */
    private Integer end;

    /**
     * 分组层级, 从 1 开始.
     */
    private Integer level;

    /**
     * 是否折叠隐藏.
     */
    private Boolean collapsed;
}
