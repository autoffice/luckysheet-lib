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

import io.github.autoffice.luckysheet.model.cell.CellData;
import io.github.autoffice.luckysheet.model.cell.CellHorizontalType;
import io.github.autoffice.luckysheet.model.cell.CellTypeEnum;
import io.github.autoffice.luckysheet.model.cell.CellVerticalType;
import io.github.autoffice.luckysheet.model.cell.TextBreakType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.Calendar;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link CellMapperToLuckySheet} 的单元格映射逻辑，覆盖各种单元格类型、样式和边界情况.
 */
class CellMapperToLuckySheetTest {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private XSSFRow row;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test");
        row = sheet.createRow(0);
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    // ========== Cell Type Tests ==========

    @Test
    void mapToCell_stringCell() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Hello World");
        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Hello World", cellData.getV().getV());
        assertEquals(CellTypeEnum.STRING, cellData.getV().getCt().getT());
    }

    @Test
    void mapToCell_numericCell() {
        XSSFCell cell = row.createCell(0, CellType.NUMERIC);
        cell.setCellValue(123.45);
        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("123.45", cellData.getV().getV());
        assertEquals(CellTypeEnum.NUMBER, cellData.getV().getCt().getT());
    }

    @Test
    void mapToCell_formulaCell() {
        XSSFCell cell = row.createCell(0, CellType.FORMULA);
        cell.setCellFormula("SUM(A1:A10)");
        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("SUM(A1:A10)", cellData.getV().getF());
        assertEquals(CellTypeEnum.GENERAL, cellData.getV().getCt().getT());
    }

    @Test
    void mapToCell_errorCell() {
        XSSFCell cell = row.createCell(0, CellType.ERROR);
        cell.setCellErrorValue((byte) 0x07);
        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellTypeEnum.E, cellData.getV().getCt().getT());
    }

    @Test
    void mapToCell_blankCell() {
        XSSFCell cell = row.createCell(0, CellType.BLANK);
        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData);
    }

    @Test
    void mapToCell_dateCell() {
        XSSFCell cell = row.createCell(0, CellType.NUMERIC);
        Calendar cal = Calendar.getInstance();
        cal.set(2025, Calendar.JANUARY, 15);
        cell.setCellValue(cal.getTime());

        XSSFCellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat((short) 14);
        cell.setCellStyle(dateStyle);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellTypeEnum.DATETIME, cellData.getV().getCt().getT());
    }

    @Test
    void mapToCell_richTextString() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        XSSFRichTextString richText = new XSSFRichTextString("Hello World");
        XSSFFont font1 = workbook.createFont();
        font1.setBold(true);
        richText.applyFont(0, 5, font1);
        cell.setCellValue(richText);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellTypeEnum.INLINESTR, cellData.getV().getCt().getT());
        assertNotNull(cellData.getV().getCt().getS());
    }

    // ========== Style Tests ==========

    @Test
    void mapToCell_nullCellStyle() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");
        cell.setCellStyle(null);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Test", cellData.getV().getV());
    }

    @Test
    void mapToCell_nullFont() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Test", cellData.getV().getV());
    }

    @Test
    void mapToCell_fontWithColor() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setColor(new XSSFColor(new byte[]{(byte) 0xFF, 0, 0}, null));
        font.setBold(true);
        font.setItalic(true);
        font.setUnderline(FontUnderline.SINGLE.getByteValue());
        font.setStrikeout(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getFc());
        assertTrue(cellData.getV().getBl().isPoiValue());
        assertTrue(cellData.getV().getIt().isPoiValue());
    }

    @Test
    void mapToCell_fontWithNullColor() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        style.setFont(font);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Test", cellData.getV().getV());
    }

    @Test
    void mapToCell_backgroundColor() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xFF, 0}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getBg());
    }

    @Test
    void mapToCell_nullBackgroundColor() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Test", cellData.getV().getV());
    }

    // ========== Alignment Tests ==========

    @Test
    void mapToCell_horizontalAlignmentLeft() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellHorizontalType.LEFT, cellData.getV().getHt());
    }

    @Test
    void mapToCell_horizontalAlignmentCenter() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellHorizontalType.CENTER, cellData.getV().getHt());
    }

    @Test
    void mapToCell_horizontalAlignmentRight() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellHorizontalType.RIGHT, cellData.getV().getHt());
    }

    @Test
    void mapToCell_horizontalAlignmentJustify() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.JUSTIFY);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getHt());
    }

    @Test
    void mapToCell_nullHorizontalAlignment() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Test", cellData.getV().getV());
    }

    @Test
    void mapToCell_verticalAlignmentTop() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.TOP);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellVerticalType.TOP, cellData.getV().getVt());
    }

    @Test
    void mapToCell_verticalAlignmentCenter() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellVerticalType.CENTER, cellData.getV().getVt());
    }

    @Test
    void mapToCell_verticalAlignmentBottom() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellVerticalType.BOTTOM, cellData.getV().getVt());
    }

    @Test
    void mapToCell_nullVerticalAlignment() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Test", cellData.getV().getV());
    }

    // ========== Text Rotation Tests ==========

    @Test
    void mapToCell_textRotation0() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setRotation((short) 0);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getTr());
    }

    @Test
    void mapToCell_textRotation90() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setRotation((short) 90);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getTr());
    }

    @Test
    void mapToCell_textRotation45() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setRotation((short) 45);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getTr());
    }

    // ========== Wrap Text Tests ==========

    @Test
    void mapToCell_wrapTextTrue() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(TextBreakType.LINE_WRAP, cellData.getV().getTb());
    }

    @Test
    void mapToCell_wrapTextFalse() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setWrapText(false);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(TextBreakType.OVERFLOW, cellData.getV().getTb());
    }

    // ========== Shrink to Fit Tests ==========

    @Test
    void mapToCell_shrinkToFitTrue() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setShrinkToFit(true);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getStf());
    }

    @Test
    void mapToCell_shrinkToFitFalse() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        XSSFCellStyle style = workbook.createCellStyle();
        style.setShrinkToFit(false);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getStf());
    }

    // ========== Comment Tests ==========

    @Test
    void mapToCell_withComment() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        org.apache.poi.ss.usermodel.ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(0);
        anchor.setCol2(1);
        anchor.setRow1(0);
        anchor.setRow2(1);
        XSSFComment comment = sheet.createDrawingPatriarch().createCellComment(anchor);
        comment.setString(new XSSFRichTextString("This is a comment"));
        comment.setVisible(true);
        cell.setCellComment(comment);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getPs());
        assertEquals("This is a comment", cellData.getV().getPs().getValue());
        assertTrue(cellData.getV().getPs().getIsshow());
    }

    @Test
    void mapToCell_withNullComment() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals("Test", cellData.getV().getV());
    }

    @Test
    void mapToCell_commentWithNullString() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Test");

        org.apache.poi.ss.usermodel.ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(0);
        anchor.setCol2(1);
        anchor.setRow1(0);
        anchor.setRow2(1);
        XSSFComment comment = sheet.createDrawingPatriarch().createCellComment(anchor);
        cell.setCellComment(comment);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getPs());
    }

    // ========== Data Format Tests ==========

    @Test
    void mapToCell_dataFormatString() {
        XSSFCell cell = row.createCell(0, CellType.NUMERIC);
        cell.setCellValue(1234.56);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setDataFormat((short) 2);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getCt().getFa());
    }

    @Test
    void mapToCell_nullDataFormatString() {
        XSSFCell cell = row.createCell(0, CellType.NUMERIC);
        cell.setCellValue(1234.56);

        XSSFCellStyle style = workbook.createCellStyle();
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertEquals(CellTypeEnum.NUMBER, cellData.getV().getCt().getT());
    }

    // ========== Rich Text Tests ==========

    @Test
    void mapToCell_richTextWithNullFont() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        XSSFRichTextString richText = new XSSFRichTextString("Hello World");
        XSSFFont font1 = workbook.createFont();
        font1.setBold(true);
        richText.applyFont(0, 5, font1);
        cell.setCellValue(richText);

        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        style.setFont(font);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getCt().getS());
    }

    @Test
    void mapToCell_richTextWithNullCellStyle() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        XSSFRichTextString richText = new XSSFRichTextString("Hello World");
        XSSFFont font1 = workbook.createFont();
        font1.setBold(true);
        richText.applyFont(0, 5, font1);
        cell.setCellValue(richText);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getCt().getS());
    }

    @Test
    void mapToCell_richTextWithFontColor() {
        XSSFCell cell = row.createCell(0, CellType.STRING);
        XSSFRichTextString richText = new XSSFRichTextString("Hello World");
        XSSFFont font1 = workbook.createFont();
        font1.setBold(true);
        font1.setColor(new XSSFColor(new byte[]{(byte) 0xFF, 0, 0}, null));
        richText.applyFont(0, 5, font1);
        cell.setCellValue(richText);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getCt().getS());
        assertTrue(cellData.getV().getCt().getS().size() > 0);
    }

    @Test
    void mapToCell_dataFormatFromDateMap() {
        XSSFCell cell = row.createCell(0, CellType.NUMERIC);
        cell.setCellValue(1234.56);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setDataFormat((short) 14);
        cell.setCellStyle(style);

        CellData cellData = LuckySheetFactory.createCellData();

        CellMapperToLuckySheet.mapToCell(cell, cellData);

        assertNotNull(cellData.getV().getCt().getFa());
    }
}

