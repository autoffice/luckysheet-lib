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

import io.github.autoffice.luckysheet.model.cell.Bold;
import io.github.autoffice.luckysheet.model.cell.Cancelline;
import io.github.autoffice.luckysheet.model.cell.CellData;
import io.github.autoffice.luckysheet.model.cell.CellHorizontalType;
import io.github.autoffice.luckysheet.model.cell.CellTypeEnum;
import io.github.autoffice.luckysheet.model.cell.CellVerticalType;
import io.github.autoffice.luckysheet.model.cell.Comment;
import io.github.autoffice.luckysheet.model.cell.FontFamily;
import io.github.autoffice.luckysheet.model.cell.InlineText;
import io.github.autoffice.luckysheet.model.cell.Italic;
import io.github.autoffice.luckysheet.model.cell.TextBreakType;
import io.github.autoffice.luckysheet.model.cell.TextRotateType;
import io.github.autoffice.luckysheet.model.cell.Underline;
import io.github.autoffice.luckysheet.util.Constant;
import io.github.autoffice.luckysheet.util.NumberUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.ArrayList;
import java.util.List;

import static io.github.autoffice.luckysheet.util.DateUtil.DATE_MAP;

public class CellMapperToLuckySheet {

    public static void mapToCell(XSSFCell cell, CellData cellData) {
        mapCellValue(cell, cellData);
        mapCellStyle(cell, cellData);
        mapComment(cell, cellData);
    }

    private static void mapComment(XSSFCell cell, CellData cellData) {
        XSSFComment cellComment = cell.getCellComment();
        if (cellComment == null) {
            return;
        }

        Comment comment = LuckySheetFactory.createComment();
        XSSFRichTextString commentString = cellComment.getString();
        if (commentString != null) {
            comment.setValue(commentString.getString());
        }

        comment.setIsshow(cellComment.isVisible());
        cellData.getV().setPs(comment);
    }

    private static void mapCellStyle(XSSFCell cell, CellData cellData) {
        XSSFCellStyle cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            return;
        }

        mapDataFormat(cellStyle, cellData);
        mapBackgroundColor(cellStyle, cellData);
        mapFont(cellStyle, cellData);
        mapHorizontalAlignment(cellStyle, cellData);
        mapVerticalAlignment(cellStyle, cellData);
        mapRotation(cellStyle, cellData);
        mapWrapText(cellStyle, cellData);
    }

    private static void mapWrapText(XSSFCellStyle cellStyle, CellData cellData) {
        if (cellStyle.getWrapText()) {
            cellData.getV().setTb(TextBreakType.LINE_WRAP);
        } else {
            cellData.getV().setTb(TextBreakType.OVERFLOW);
        }
    }

    private static void mapDataFormat(XSSFCellStyle cellStyle, CellData cellData) {
        String dataFormatString = cellStyle.getDataFormatString();
        if (dataFormatString == null) {
            if (DATE_MAP.get(cellStyle.getDataFormat()) != null) {
                dataFormatString = DATE_MAP.get(cellStyle.getDataFormat());
                cellData.getV().getCt().setT(CellTypeEnum.DATETIME);
            } else {
                return;
            }
        }

        cellData.getV().getCt().setFa(dataFormatString);
    }

    private static void mapRotation(XSSFCellStyle cellStyle, CellData cellData) {
        cellData.getV().setTr(TextRotateType.of(cellStyle.getRotation()));
    }

    private static void mapVerticalAlignment(XSSFCellStyle cellStyle, CellData cellData) {
        VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
        if (verticalAlignment == null) {
            return;
        }

        cellData.getV().setVt(CellVerticalType.of(verticalAlignment));
    }

    private static void mapHorizontalAlignment(XSSFCellStyle cellStyle, CellData cellData) {
        HorizontalAlignment alignment = cellStyle.getAlignment();
        if (alignment == null) {
            return;
        }

        cellData.getV().setHt(CellHorizontalType.of(alignment));
    }

    private static void mapFont(XSSFCellStyle cellStyle, CellData cellData) {
        XSSFFont font = cellStyle.getFont();
        if (font == null) {
            return;
        }

        cellData.getV().setFf(FontFamily.of(font.getFontName()));
        cellData.getV().setBl(Bold.of(font.getBold()));
        cellData.getV().setIt(Italic.of(font.getItalic()));
        cellData.getV().setUn(Underline.of(FontUnderline.valueOf(font.getUnderline())));
        cellData.getV().setCl(Cancelline.of(font.getStrikeout()));
        cellData.getV().setFs(font.getFontHeightInPoints());

        if (font.getXSSFColor() != null) {
            cellData.getV().setFc(NumberUtil.rgbToColorString(font.getXSSFColor().getRGB()));
        }
    }

    private static void mapBackgroundColor(XSSFCellStyle cellStyle, CellData cellData) {
        XSSFColor fillForegroundXSSFColor = cellStyle.getFillForegroundXSSFColor();
        if (fillForegroundXSSFColor == null) {
            return;
        }

        cellData.getV().setBg(NumberUtil.rgbToColorString(fillForegroundXSSFColor.getRGB()));
    }

    private static void mapCellValue(XSSFCell cell, CellData cellData) {
        CellType cellType = cell.getCellType();
        DataFormatter dataFormatter = new DataFormatter();
        if (cellType == CellType.STRING) {
            XSSFRichTextString richTextString = cell.getRichStringCellValue();
            if (richTextString.hasFormatting()) {
                List<InlineText> inlineTexts = new ArrayList<>();
                mapRichTextString(richTextString, cell.getCellStyle(), inlineTexts);
                cellData.getV().getCt().setS(inlineTexts);
                cellData.getV().getCt().setFa(Constant.FA_GENERAL);
                cellData.getV().getCt().setT(CellTypeEnum.INLINESTR);
            } else {
                cellData.getV().setV(cell.getStringCellValue());
                cellData.getV().setM(dataFormatter.formatCellValue(cell));
                cellData.getV().getCt().setFa(Constant.FA_GENERAL);
                cellData.getV().getCt().setT(CellTypeEnum.STRING);
            }
        } else if (cellType == CellType.FORMULA) {
            cellData.getV().setF(cell.getCellFormula());
            cellData.getV().setM(dataFormatter.formatCellValue(cell));
            cellData.getV().getCt().setFa(Constant.FA_GENERAL);
            cellData.getV().getCt().setT(CellTypeEnum.GENERAL);
        } else if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                cellData.getV().setV(String.valueOf(cell.getNumericCellValue()));
                cellData.getV().setM(dataFormatter.formatCellValue(cell));
                cellData.getV().getCt().setT(CellTypeEnum.DATETIME);
            } else {
                cellData.getV().setV(String.valueOf(cell.getNumericCellValue()));
                cellData.getV().setM(dataFormatter.formatCellValue(cell));
                cellData.getV().getCt().setT(CellTypeEnum.NUMBER);
            }
        } else if (cellType == CellType.BOOLEAN) {
            cellData.getV().setV(String.valueOf(cell.getNumericCellValue()));
            cellData.getV().setM(dataFormatter.formatCellValue(cell));
            cellData.getV().getCt().setT(CellTypeEnum.B);
        } else if (cellType == CellType.ERROR) {
            cellData.getV().setV(cell.getErrorCellString());
            cellData.getV().setM(dataFormatter.formatCellValue(cell));
            cellData.getV().getCt().setT(CellTypeEnum.E);
        }
    }

    private static void mapRichTextString(XSSFRichTextString richTextString, XSSFCellStyle cellStyle, List<InlineText> inlineTexts) {
        String string = richTextString.getString();
        for (int i = 0; i < richTextString.numFormattingRuns(); i++) {
            int lengthOfFormattingRun = richTextString.getLengthOfFormattingRun(i);
            int indexOfFormattingRun = richTextString.getIndexOfFormattingRun(i);
            String substring = string.substring(indexOfFormattingRun, indexOfFormattingRun + lengthOfFormattingRun);
            XSSFFont font = richTextString.getFontOfFormattingRun(i);

            if (font == null && cellStyle != null) {
                font = cellStyle.getFont();
            }

            InlineText inlineText = LuckySheetFactory.createInlineText();
            if (font != null) {
                inlineText.setFf(FontFamily.of(font.getFontName()));
                inlineText.setBl(Bold.of(font.getBold()));
                inlineText.setIt(Italic.of(font.getItalic()));
                inlineText.setUn(Underline.of(FontUnderline.valueOf(font.getUnderline())));
                inlineText.setCl(Cancelline.of(font.getStrikeout()));
                inlineText.setFs(font.getFontHeightInPoints());

                if (font.getXSSFColor() != null) {
                    inlineText.setFc(NumberUtil.rgbToColorString(font.getXSSFColor().getRGB()));
                }
            }

            inlineText.setV(substring);
            inlineTexts.add(inlineText);
        }
    }
}
