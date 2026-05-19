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
package io.github.autoffice.luckysheet.mapper;

import io.github.autoffice.luckysheet.model.LuckyFile;
import io.github.autoffice.luckysheet.model.LuckyFileInfo;
import io.github.autoffice.luckysheet.model.cell.CellData;
import io.github.autoffice.luckysheet.model.cell.CellType;
import io.github.autoffice.luckysheet.model.cell.CellValue;
import io.github.autoffice.luckysheet.model.cell.Comment;
import io.github.autoffice.luckysheet.model.cell.InlineText;
import io.github.autoffice.luckysheet.model.image.ImageBorder;
import io.github.autoffice.luckysheet.model.image.ImageCrop;
import io.github.autoffice.luckysheet.model.image.ImagePosition;
import io.github.autoffice.luckysheet.model.image.SheetImage;
import io.github.autoffice.luckysheet.model.sheet.Border;
import io.github.autoffice.luckysheet.model.sheet.BorderRangeType;
import io.github.autoffice.luckysheet.model.sheet.BorderStyleType;
import io.github.autoffice.luckysheet.model.sheet.Frozen;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.util.NumberUtil;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Luckysheet 模型对象工厂.
 *
 * <p>集中创建 {@link LuckyFile}、{@link LuckySheet}、{@link CellData} 等模型实例,
 * 确保对象初始化时具备必要的默认值.</p>
 */
public class LuckySheetFactory {
    /**
     * 创建空的 LuckyFile 实例.
     *
     * @return 新的 LuckyFile 实例
     */
    public static LuckyFile createLuckyFile() {
        return new LuckyFile();
    }

    /**
     * 创建空的 LuckyFileInfo 实例.
     *
     * @return 新的 LuckyFileInfo 实例
     */
    public static LuckyFileInfo createLuckyFileInfo() {
        return new LuckyFileInfo();
    }

    /**
     * 创建 LuckySheet 实例并设置工作表名称.
     *
     * @param sheetName 工作表名称
     * @return 初始化后的 LuckySheet 实例
     */
    public static LuckySheet createLuckySheet(String sheetName) {
        LuckySheet luckySheet = new LuckySheet();
        luckySheet.setName(sheetName);
        return luckySheet;
    }

    /**
     * 创建 CellData 实例, 内含初始化的 CellValue 和 CellType.
     *
     * @return 初始化后的 CellData 实例
     */
    public static CellData createCellData() {
        CellData cellData = new CellData();
        CellValue v = new CellValue();
        v.setCt(new CellType());
        cellData.setV(v);

        return cellData;
    }

    /**
     * 创建空的 InlineText 实例.
     *
     * @return 新的 InlineText 实例
     */
    public static InlineText createInlineText() {
        return new InlineText();
    }

    /**
     * 创建空的 Comment 实例.
     *
     * @return 新的 Comment 实例
     */
    public static Comment createComment() {
        return new Comment();
    }

    /**
     * 创建 Border 实例, 默认范围类型为 CELL.
     *
     * @return 初始化后的 Border 实例
     */
    public static Border createBorder() {
        Border border = new Border();
        border.setRangeType(BorderRangeType.CELL);
        border.setValue(new Border.Value());
        return border;
    }

    /**
     * 从 POI 边框颜色和样式创建 Border.Style.
     *
     * @param borderColor POI 边框颜色
     * @param borderStyle POI 边框样式
     * @return Border.Style 实例, 颜色无效时返回 null
     */
    public static Border.Style createBorderStyle(XSSFColor borderColor, BorderStyle borderStyle) {
        if (borderColor == null) {
            return null;
        }

        Border.Style style = new Border.Style();
        if (borderColor.isAuto()) {
            style.setColor("rgb(0,0,0)");
        } else {
            byte[] rgb = borderColor.getRGBWithTint();
            if (rgb == null || rgb.length == 0) {
                rgb = borderColor.getRGB();
            }
            if (rgb == null || rgb.length == 0) {
                style.setColor("rgb(0,0,0)");
            } else {
                style.setColor(NumberUtil.rgbToColorString(rgb));
            }
        }

        style.setStyle(BorderStyleType.of(borderStyle));
        return style;
    }

    /**
     * 创建 SheetImage 实例, 内含初始化的 border、crop 和 position.
     *
     * @return 初始化后的 SheetImage 实例
     */
    public static SheetImage createImage() {
        SheetImage sheetImage = new SheetImage();
        sheetImage.setBorder(new ImageBorder());
        sheetImage.setCrop(new ImageCrop());
        sheetImage.setPosition(new ImagePosition());
        return sheetImage;
    }

    /**
     * 判断 Border 是否包含至少一条有效边框样式 (上/右/下/左).
     *
     * @param border 待检查的 Border 实例
     * @return 包含有效样式返回 true, 否则返回 false
     */
    public static boolean hasBorderStyle(Border border) {
        return border.getValue().getT() != null
                || border.getValue().getR() != null
                || border.getValue().getB() != null
                || border.getValue().getL() != null;
    }

    /**
     * 创建指定行列位置的 Border 实例.
     *
     * @param row 行索引
     * @param col 列索引
     * @return 初始化后的 Border 实例
     */
    public static Border createBorder(int row, int col) {
        Border borderTmp = new Border();
        borderTmp.setRangeType(BorderRangeType.CELL);
        Border.Value value = new Border.Value();
        value.setRow_index(row);
        value.setCol_index(col);
        borderTmp.setValue(value);
        return borderTmp;
    }

    /**
     * 从 Luckysheet Border 的 style 和 color 属性创建 Border.Style.
     *
     * @param border 源 Border 实例
     * @return 新的 Border.Style 实例
     */
    public static Border.Style createBorderStyle(Border border) {
        Border.Style style = new Border.Style();
        style.setStyle(border.getStyle());
        style.setColor(border.getColor());
        return style;
    }

    /**
     * 创建空的 Frozen 实例.
     *
     * @return 新的 Frozen 实例
     */
    public static Frozen createFrozen() {
        return new Frozen();
    }
}
