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
import java.util.Map;

/**
 * Luckysheet 图表描述 (有损转换).
 *
 * <p>Excel 侧 XSSFChart 与 Luckysheet 前端 echarts 配置模型差异极大, 本实现在
 * 双向转换时仅保留图表的类型、位置、数据范围和标题等共通字段. 样式、主题、交互
 * 参数等信息不保真.</p>
 */
@Data
public class SheetChart {
    /**
     * 图表唯一 ID.
     */
    private String chart_id;

    /**
     * 图表宽度 (像素).
     */
    private Integer width;

    /**
     * 图表高度 (像素).
     */
    private Integer height;

    /**
     * 图表左上角相对 sheet 的 x 偏移 (像素).
     */
    private Integer left;

    /**
     * 图表左上角相对 sheet 的 y 偏移 (像素).
     */
    private Integer top;

    /**
     * 所属 sheet 索引.
     */
    private String sheetIndex;

    /**
     * 是否显示范围框.
     */
    private Boolean needRangeShow;

    /**
     * 图表详细配置, 如类型、标题、数据范围等.
     */
    private ChartOptions chartOptions;

    /**
     * 图表详细配置子对象.
     */
    @Data
    public static class ChartOptions {
        private String chart_id;
        private String chartAllType;
        private String chartPro;
        /**
         * 图表类型: bar / column / line / pie ...
         */
        private String chartType;
        private String chartStyle;
        /**
         * 数据范围列表, 每个元素指定一个 Range.
         */
        private List<Range> rangeArray;
        /**
         * 标题.
         */
        private String title;
        /**
         * 展平后的二维数据, 便于前端直接渲染.
         */
        private List<List<Object>> chartData;
        /**
         * echarts 原始 option (透传).
         */
        private Map<String, Object> defaultOption;
    }
}
