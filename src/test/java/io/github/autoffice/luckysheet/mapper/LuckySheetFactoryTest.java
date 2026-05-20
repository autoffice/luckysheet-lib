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
import io.github.autoffice.luckysheet.model.cell.Comment;
import io.github.autoffice.luckysheet.model.cell.InlineText;
import io.github.autoffice.luckysheet.model.image.SheetImage;
import io.github.autoffice.luckysheet.model.sheet.Border;
import io.github.autoffice.luckysheet.model.sheet.BorderRangeType;
import io.github.autoffice.luckysheet.model.sheet.BorderStyleType;
import io.github.autoffice.luckysheet.model.sheet.Frozen;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link LuckySheetFactory} 的所有工厂方法, 涵盖边框样式、颜色处理等分支.
 */
class LuckySheetFactoryTest {

    // ========== createLuckyFile ==========

    @Test
    void createLuckyFile_returnsNonNull() {
        LuckyFile file = LuckySheetFactory.createLuckyFile();
        assertNotNull(file);
    }

    // ========== createLuckyFileInfo ==========

    @Test
    void createLuckyFileInfo_returnsNonNull() {
        LuckyFileInfo info = LuckySheetFactory.createLuckyFileInfo();
        assertNotNull(info);
    }

    // ========== createLuckySheet ==========

    @Test
    void createLuckySheet_setsName() {
        LuckySheet sheet = LuckySheetFactory.createLuckySheet("TestSheet");
        assertNotNull(sheet);
        assertEquals("TestSheet", sheet.getName());
    }

    // ========== createCellData ==========

    @Test
    void createCellData_initializesValueAndType() {
        CellData cell = LuckySheetFactory.createCellData();
        assertNotNull(cell);
        assertNotNull(cell.getV());
        assertNotNull(cell.getV().getCt());
    }

    // ========== createInlineText ==========

    @Test
    void createInlineText_returnsNonNull() {
        InlineText text = LuckySheetFactory.createInlineText();
        assertNotNull(text);
    }

    // ========== createComment ==========

    @Test
    void createComment_returnsNonNull() {
        Comment comment = LuckySheetFactory.createComment();
        assertNotNull(comment);
    }

    // ========== createBorder (no args) ==========

    @Test
    void createBorder_noArgs_initializesWithCellRangeType() {
        Border border = LuckySheetFactory.createBorder();
        assertNotNull(border);
        assertEquals(BorderRangeType.CELL, border.getRangeType());
        assertNotNull(border.getValue());
    }

    // ========== createBorder (row, col) ==========

    @Test
    void createBorder_withRowCol_setsIndices() {
        Border border = LuckySheetFactory.createBorder(5, 10);
        assertNotNull(border);
        assertEquals(BorderRangeType.CELL, border.getRangeType());
        assertNotNull(border.getValue());
        assertEquals(5, border.getValue().getRow_index());
        assertEquals(10, border.getValue().getCol_index());
    }

    // ========== createBorderStyle (XSSFColor, BorderStyle) ==========

    @Test
    void createBorderStyle_nullColor_returnsNull() {
        Border.Style style = LuckySheetFactory.createBorderStyle(null, BorderStyle.THIN);
        assertNull(style);
    }

    @Test
    void createBorderStyle_autoColor_setsBlack() {
        XSSFColor color = new XSSFColor();
        color.setAuto(true);
        Border.Style style = LuckySheetFactory.createBorderStyle(color, BorderStyle.THIN);
        assertNotNull(style);
        assertEquals("rgb(0,0,0)", style.getColor());
        assertEquals(BorderStyleType.THIN, style.getStyle());
    }

    @Test
    void createBorderStyle_validRgbWithTint_usesRgbWithTint() {
        byte[] rgb = new byte[]{(byte) 255, 0, 0};
        XSSFColor color = new XSSFColor(rgb, null);
        Border.Style style = LuckySheetFactory.createBorderStyle(color, BorderStyle.MEDIUM);
        assertNotNull(style);
        assertTrue(style.getColor().startsWith("rgb("));
        assertEquals(BorderStyleType.MEDIUM, style.getStyle());
    }

    @Test
    void createBorderStyle_rgbWithTintReturnsNull_fallsBackToGetRgb() {
        // XSSFColor with indexed color may return null for getRGBWithTint
        XSSFColor color = new XSSFColor();
        color.setIndexed(1);
        Border.Style style = LuckySheetFactory.createBorderStyle(color, BorderStyle.DASHED);
        assertNotNull(style);
        // Falls back to getRGB or default black
        assertNotNull(style.getColor());
    }

    @Test
    void createBorderStyle_bothRgbMethodsReturnNull_defaultsToBlack() {
        // Create a color that returns null/empty for both getRGBWithTint and getRGB
        XSSFColor color = new XSSFColor();
        Border.Style style = LuckySheetFactory.createBorderStyle(color, BorderStyle.DOTTED);
        assertNotNull(style);
        assertEquals("rgb(0,0,0)", style.getColor());
        assertEquals(BorderStyleType.DOTTED, style.getStyle());
    }

    // ========== createImage ==========

    @Test
    void createImage_initializesAllComponents() {
        SheetImage image = LuckySheetFactory.createImage();
        assertNotNull(image);
        assertNotNull(image.getBorder());
        assertNotNull(image.getCrop());
        assertNotNull(image.getPosition());
    }

    // ========== hasBorderStyle ==========

    @Test
    void hasBorderStyle_allSidesNull_returnsFalse() {
        Border border = LuckySheetFactory.createBorder();
        assertFalse(LuckySheetFactory.hasBorderStyle(border));
    }

    @Test
    void hasBorderStyle_topPresent_returnsTrue() {
        Border border = LuckySheetFactory.createBorder();
        Border.Style style = new Border.Style();
        border.getValue().setT(style);
        assertTrue(LuckySheetFactory.hasBorderStyle(border));
    }

    @Test
    void hasBorderStyle_rightPresent_returnsTrue() {
        Border border = LuckySheetFactory.createBorder();
        Border.Style style = new Border.Style();
        border.getValue().setR(style);
        assertTrue(LuckySheetFactory.hasBorderStyle(border));
    }

    @Test
    void hasBorderStyle_bottomPresent_returnsTrue() {
        Border border = LuckySheetFactory.createBorder();
        Border.Style style = new Border.Style();
        border.getValue().setB(style);
        assertTrue(LuckySheetFactory.hasBorderStyle(border));
    }

    @Test
    void hasBorderStyle_leftPresent_returnsTrue() {
        Border border = LuckySheetFactory.createBorder();
        Border.Style style = new Border.Style();
        border.getValue().setL(style);
        assertTrue(LuckySheetFactory.hasBorderStyle(border));
    }

    // ========== createBorderStyle (Border) ==========

    @Test
    void createBorderStyle_fromBorder_copiesStyleAndColor() {
        Border source = new Border();
        source.setStyle(BorderStyleType.THICK);
        source.setColor("rgb(255,0,0)");
        Border.Style style = LuckySheetFactory.createBorderStyle(source);
        assertNotNull(style);
        assertEquals(BorderStyleType.THICK, style.getStyle());
        assertEquals("rgb(255,0,0)", style.getColor());
    }

    // ========== createFrozen ==========

    @Test
    void createFrozen_returnsNonNull() {
        Frozen frozen = LuckySheetFactory.createFrozen();
        assertNotNull(frozen);
    }
}
