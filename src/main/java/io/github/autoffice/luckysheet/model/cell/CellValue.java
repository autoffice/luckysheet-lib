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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Map;

/**
 * @see <a
 * href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC">
 * 单元格</a>
 */
@Data
public class CellValue {

    /**
     * 背景颜色Hex. eg: {@code #fff000}
     */
    private String bg;

    /**
     * 粗体. 0 常规 、 1加粗
     */
    private Bold bl;

    /**
     * 斜体. 0 常规 、 1 斜体
     */
    private Italic it;

    /**
     * 字体
     */
    private FontFamily ff;

    /**
     * 字体大小
     */
    private Short fs;

    /**
     * 字体颜色. #fff000
     */
    private String fc;

    /**
     * 水平对齐. 0 居中、1 左、2右
     */
    private CellHorizontalType ht;

    /**
     * 垂直对齐. 0,居中; 1,上; 2,下
     */
    private CellVerticalType vt;

    /**
     * 原始值.可能为字符串类型,也可能为数字类型.
     * <p>
     * 理论上来说,设定好{@link CellValue#ct}后,此处为string不影响显示
     * <p>
     */
    private String v;

    /**
     * 单元格值格式：文本、时间等
     */
    private CellType ct;

    /**
     * 显示值
     */
    private String m;

    /**
     * 公式. eg: {@code "=SUM(233)" }
     */
    private String f;


    /**
     * 删除线. 0 常规 、 1 删除线
     */
    private Cancelline cl;

    /**
     * 下划线. 0 无 、 1 有
     */
    private Underline un;

    /**
     * 合并单元格
     */
    private MergeCell mc;
    /**
     * 竖排文字,<b>参考微软excel中 开始->对齐方式->方向</b>
     * <p>Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90
     */
    private TextRotateType tr;
    /**
     * 文字旋转角度,取值范围[0-180].
     *
     * @see XSSFCellStyle#getRotation()
     */
    private Short rt;
    /**
     * 文本换行. 0 截断、1溢出、2 自动换行
     */
    private TextBreakType tb;

    /**
     * 批注
     */
    private Comment ps;

    /**
     * TODO 未知属性
     */
    private Integer qp;
    private Map<String, Object> customKey;
}
