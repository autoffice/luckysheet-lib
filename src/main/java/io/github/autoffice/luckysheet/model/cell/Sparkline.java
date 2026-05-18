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

import java.util.List;
import java.util.Map;

/**
 * 迷你图 (Sparkline) 的序列化载体.
 *
 * <p>Luckysheet 中迷你图以 cell.v.spl 的形式序列化, 内部包含 shapes 等渲染产物.
 * Excel 侧通过 xlsx 扩展列表存储 sparkline group; 本模型主要保留 Luckysheet 一侧
 * 字段以便 round-trip 时避免信息丢失. Excel → Luckysheet 时会填充 dataRange 与 type;
 * Luckysheet → Excel 时以 dataRange 为输入生成最基本的 sparkline group.</p>
 */
@Data
public class Sparkline {
    /**
     * 迷你图类型: line / column / winloss.
     */
    private String type;

    /**
     * 数据源区域 (A1 风格), eg: {@code Sheet1!A1:A5}
     */
    private String dataRange;

    /**
     * Luckysheet 渲染产物 (形状信息).
     */
    private Map<String, Object> shapes;

    /**
     * Luckysheet 渲染产物形状顺序.
     */
    private List<Integer> shapeseq;

    private Integer offsetX;
    private Integer offsetY;
    private Integer pixelWidth;
    private Integer pixelHeight;
}
