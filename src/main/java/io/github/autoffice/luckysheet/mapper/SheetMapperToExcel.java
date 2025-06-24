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
import io.github.autoffice.luckysheet.model.cell.MergeCell;
import io.github.autoffice.luckysheet.model.sheet.BoolStatus;
import io.github.autoffice.luckysheet.model.sheet.Border;
import io.github.autoffice.luckysheet.model.sheet.BorderRangeType;
import io.github.autoffice.luckysheet.model.sheet.BorderStyleType;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import io.github.autoffice.luckysheet.util.NumberUtil;
import io.github.autoffice.luckysheet.util.Util;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class SheetMapperToExcel {

    public static final Short LUCKY_SHEET_DEFAULT_ROW_HEIGHT_IN_PIXEL = 19;
    public static final Short LUCY_SHEET_DEFAULT_COL_WIDTH_IN_PIXEL = 73;

    public static void mapToSheet(LuckySheet luckySheet, XSSFSheet sheet) {
        mapDefaultRowHeight(luckySheet.getDefaultRowHeight(), sheet);
        mapDefaultColumnWidth(luckySheet.getDefaultColWidth(), sheet);

        for (CellData cellData : luckySheet.getCelldata()) {
            XSSFCell cell = PoiFactory.createOrGetCell(sheet, cellData.getR(), cellData.getC());
            CellMapperToExcel.mapToCell(cellData, cell);
        }

        mapSheetStatus(luckySheet.getStatus(), sheet);
        mapSheetHidden(luckySheet.getHide(), sheet);
        mapMergeCell(luckySheet.getConfig().getMerge(), sheet);
        mapRowHeight(luckySheet.getConfig().getRowlen(), sheet);
        mapColumnWith(luckySheet.getConfig().getColumnlen(), sheet);
        mapRowHidden(luckySheet.getConfig().getRowhidden(), sheet);
        mapColumnHidden(luckySheet.getConfig().getColhidden(), sheet);
        mapBorder(luckySheet.getConfig().getBorderInfo(), sheet);
        mapGridLines(luckySheet.getShowGridLines(), sheet);

        ImageMapperToExcel.mapToSheet(luckySheet.getImages(), sheet);
    }

    private static void mapGridLines(BoolStatus showGridLines, XSSFSheet sheet) {
        if (showGridLines == null) {
            return;
        }

        sheet.setDisplayGridlines(showGridLines.isPoiValue());
    }

    private static void mapBorder(List<Border> borderList, XSSFSheet sheet) {
        if (CollectionUtils.isEmpty(borderList)) {
            return;
        }

        ArrayList<Border> borderListTmp = new ArrayList<>();
        for (Border border : borderList) {
            if (border.getRangeType() == BorderRangeType.RANGE) {
                borderListTmp.addAll(mapRangeBorderToCellBorder(border));
            }
        }
        borderList.addAll(borderListTmp);

        for (Border border : borderList) {
            if (border.getRangeType() == BorderRangeType.CELL && LuckySheetFactory.hasBorderStyle(border)) {
                XSSFCell cell = PoiFactory.createOrGetCell(sheet, border.getValue().getRow_index(), border.getValue().getCol_index());
                mapCellBorder(cell, border.getValue());
            }
        }
    }

    private static List<Border> mapRangeBorderToCellBorder(Border border) {
        List<Border> borders = new ArrayList<>();
        if (CollectionUtils.isEmpty(border.getRange())) {
            return borders;
        }

        for (Range range : border.getRange()) {
            if (CollectionUtils.size(range.getRow()) != 2 || CollectionUtils.size(range.getColumn()) != 2) {
                continue;
            }

            final Integer rowTop = range.getRow().get(0);
            final Integer rowBottom = range.getRow().get(1);
            final Integer colLeft = range.getColumn().get(0);
            final Integer colRight = range.getColumn().get(1);
            final Border.Style borderStyle = LuckySheetFactory.createBorderStyle(border);
            switch (border.getBorderType()) {
                case LEFT: {
                    for (int row = rowTop; row <= rowBottom; row++) {
                        Border borderTmp = LuckySheetFactory.createBorder(row, colLeft);
                        borderTmp.getValue().setL(borderStyle);
                        borders.add(borderTmp);
                    }
                    break;
                }
                case RIGHT:{
                    for (int row = rowTop; row <= rowBottom; row++) {
                        Border borderTmp = LuckySheetFactory.createBorder(row, colRight);
                        borderTmp.getValue().setR(borderStyle);
                        borders.add(borderTmp);
                    }
                    break;
                }
                case TOP: {
                    for (int col = colLeft; col <= colRight; col++) {
                        Border borderTmp = LuckySheetFactory.createBorder(rowTop, col);
                        borderTmp.getValue().setT(borderStyle);
                        borders.add(borderTmp);
                    }
                    break;
                }
                case BOTTOM: {
                    for (int col = colLeft; col <= colRight; col++) {
                        Border borderTmp = LuckySheetFactory.createBorder(rowBottom, col);
                        borderTmp.getValue().setB(borderStyle);
                        borders.add(borderTmp);
                    }
                    break;
                }
                case ALL: {
                    for (int row = rowTop; row <= rowBottom; row++) {
                        for (int col = colLeft; col <= colRight; col++) {
                            Border borderTmp = LuckySheetFactory.createBorder(row, col);
                            borderTmp.getValue().setL(borderStyle);
                            borderTmp.getValue().setR(borderStyle);
                            borderTmp.getValue().setT(borderStyle);
                            borderTmp.getValue().setB(borderStyle);
                            borders.add(borderTmp);
                        }
                    }
                    break;
                }
                case OUTSIDE: {
                    for (int row = rowTop; row <= rowBottom; row++) {
                        for (int col = colLeft; col <= colRight; col++) {
                            Border borderTmp = LuckySheetFactory.createBorder(row, col);
                            if (row == rowTop) {
                                borderTmp.getValue().setT(borderStyle);
                            }

                            if (row == rowBottom) {
                                borderTmp.getValue().setB(borderStyle);
                            }

                            if (col == colLeft) {
                                borderTmp.getValue().setL(borderStyle);
                            }

                            if (col == colRight) {
                                borderTmp.getValue().setR(borderStyle);
                            }

                            if (LuckySheetFactory.hasBorderStyle(borderTmp)) {
                                borders.add(borderTmp);
                            }
                        }
                    }
                    break;
                }
                case INSIDE : {
                    for (int row = rowTop; row <= rowBottom; row++) {
                        for (int col = colLeft; col <= colRight; col++) {
                            Border borderTmp = LuckySheetFactory.createBorder(row, col);
                            borderTmp.getValue().setT(borderStyle);
                            borderTmp.getValue().setB(borderStyle);
                            borderTmp.getValue().setL(borderStyle);
                            borderTmp.getValue().setR(borderStyle);

                            if (row == rowTop) {
                                borderTmp.getValue().setT(null);
                            }

                            if (row == rowBottom) {
                                borderTmp.getValue().setB(null);
                            }

                            if (col == colLeft) {
                                borderTmp.getValue().setL(null);
                            }

                            if (col == colRight) {
                                borderTmp.getValue().setR(null);
                            }

                            if (LuckySheetFactory.hasBorderStyle(borderTmp)) {
                                borders.add(borderTmp);
                            }
                        }
                    }
                    break;
                }
                case HORIZONTAL: {
                    for (int row = rowTop; row <= rowBottom; row++) {
                        for (int col = colLeft; col <= colRight; col++) {
                            Border borderTmp = LuckySheetFactory.createBorder(row, col);
                            if (row != rowBottom) {
                                borderTmp.getValue().setB(borderStyle);
                            }

                            if (LuckySheetFactory.hasBorderStyle(borderTmp)) {
                                borders.add(borderTmp);
                            }
                        }
                    }
                    break;
                }
                case VERTICAL: {
                    for (int row = rowTop; row <= rowBottom; row++) {
                        for (int col = colLeft; col <= colRight; col++) {
                            Border borderTmp = LuckySheetFactory.createBorder(row, col);
                            if (col != colRight) {
                                borderTmp.getValue().setR(borderStyle);
                            }

                            if (LuckySheetFactory.hasBorderStyle(borderTmp)) {
                                borders.add(borderTmp);
                            }
                        }
                    }
                    break;
                }
                case NONE: {
                    for (int row = rowTop; row <= rowBottom; row++) {
                        for (int col = colLeft; col <= colRight; col++) {
                            Border borderTmp = LuckySheetFactory.createBorder(row, col);
                            borderStyle.setStyle(BorderStyleType.NONE);
                            borderTmp.getValue().setL(borderStyle);
                            borderTmp.getValue().setR(borderStyle);
                            borderTmp.getValue().setT(borderStyle);
                            borderTmp.getValue().setB(borderStyle);
                            borders.add(borderTmp);
                        }
                    }
                    break;
                }
            }
        }

        return borders;
    }

    private static void mapCellBorder(XSSFCell cell, Border.Value value) {
        XSSFCellStyle cellStyle = cell.getCellStyle();

        if (value.getT() != null) {
            cellStyle.setBorderTop(value.getT().getStyle().getPoiValue());
            cellStyle.setTopBorderColor(new XSSFColor(NumberUtil.colorStringToRgb(value.getT().getColor()), null));
        }

        if (value.getR() != null) {
            cellStyle.setBorderRight(value.getR().getStyle().getPoiValue());
            cellStyle.setRightBorderColor(new XSSFColor(NumberUtil.colorStringToRgb(value.getR().getColor()), null));
        }

        if (value.getB() != null) {
            cellStyle.setBorderBottom(value.getB().getStyle().getPoiValue());
            cellStyle.setBottomBorderColor(new XSSFColor(NumberUtil.colorStringToRgb(value.getB().getColor()), null));
        }

        if (value.getL() != null) {
            cellStyle.setBorderLeft(value.getL().getStyle().getPoiValue());
            cellStyle.setLeftBorderColor(new XSSFColor(NumberUtil.colorStringToRgb(value.getL().getColor()), null));
        }
    }

    private static void mapColumnHidden(Map<Integer, Integer> colhidden, XSSFSheet sheet) {
        if (MapUtils.isEmpty(colhidden)) {
            return;
        }

        for (Integer col : colhidden.keySet()) {
            sheet.setColumnHidden(col, true);
        }
    }

    private static void mapRowHidden(Map<Integer, Integer> rowhidden, XSSFSheet sheet) {
        if (MapUtils.isEmpty(rowhidden)) {
            return;
        }

        for (Integer rowNum : rowhidden.keySet()) {
            XSSFRow row = PoiFactory.createOrGetRow(sheet, rowNum);
            row.setZeroHeight(true);
        }
    }

    private static void mapColumnWith(Map<Integer, Integer> columnlen, XSSFSheet sheet) {
        if (MapUtils.isEmpty(columnlen)) {
            return;
        }

        for (Map.Entry<Integer, Integer> entry : columnlen.entrySet()) {
            Integer colNum = entry.getKey();
            Integer width = entry.getValue();
            sheet.setColumnWidth(colNum, NumberUtil.pixel2PoiColWidth(width));
        }
    }

    /**
     * 1pt = 20twips
     * 1px = 15twip
     * 1px = 0.75pt
     */
    private static void mapRowHeight(Map<Integer, Integer> rowlen, XSSFSheet sheet) {
        if (MapUtils.isEmpty(rowlen)) {
            return;
        }

        for (Map.Entry<Integer, Integer> entry : rowlen.entrySet()) {
            Integer rowNum = entry.getKey();
            Integer height = entry.getValue();
            Row row = PoiFactory.createOrGetRow(sheet, rowNum);
            row.setHeight((short) NumberUtil.pixel2Twips(height));
        }
    }

    private static void mapMergeCell(Map<String, MergeCell> merge, XSSFSheet sheet) {
        if (MapUtils.isEmpty(merge)) {
            return;
        }

        for (MergeCell value : merge.values()) {
            CellRangeAddress cellAddresses = new CellRangeAddress(value.getR(),
                    value.getR() + value.getRs() - 1,
                    value.getC(),
                    value.getC() + value.getCs() - 1);
            sheet.addMergedRegion(cellAddresses);
        }
    }

    private static void mapSheetHidden(BoolStatus hide, XSSFSheet sheet) {
        if (hide == null) {
            return;
        }

        sheet.getWorkbook().setSheetHidden(sheet.getWorkbook().getSheetIndex(sheet), hide.isPoiValue());
    }

    private static void mapSheetStatus(BoolStatus status, XSSFSheet sheet) {
        if (status == null) {
            return;
        }

        sheet.setSelected(status.isPoiValue());
    }

    private static void mapDefaultColumnWidth(Short defaultColumnWidth, XSSFSheet sheet) {
        sheet.setDefaultColumnWidth(NumberUtil.pixel2CharacterLen(Util.requireNonNullElse(defaultColumnWidth, LUCY_SHEET_DEFAULT_COL_WIDTH_IN_PIXEL)));
    }

    private static void mapDefaultRowHeight(Short defaultRowHeight, XSSFSheet sheet) {
        sheet.setDefaultRowHeight((short) NumberUtil.pixel2Twips(Util.requireNonNullElse(defaultRowHeight, LUCKY_SHEET_DEFAULT_ROW_HEIGHT_IN_PIXEL)));
    }
}
