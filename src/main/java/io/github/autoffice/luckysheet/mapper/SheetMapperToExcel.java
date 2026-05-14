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
import io.github.autoffice.luckysheet.model.sheet.DataVerification;
import io.github.autoffice.luckysheet.model.sheet.BorderRangeType;
import io.github.autoffice.luckysheet.model.sheet.BorderStyleType;
import io.github.autoffice.luckysheet.model.sheet.Frozen;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import io.github.autoffice.luckysheet.util.NumberUtil;
import io.github.autoffice.luckysheet.util.Util;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
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
        mapFrozen(luckySheet.getFrozen(), sheet);
        mapTabColor(luckySheet.getColor(), sheet);
        mapDataVerification(luckySheet.getDataVerification(), sheet);

        ImageMapperToExcel.mapToSheet(luckySheet.getImages(), sheet);
    }

    private static void mapTabColor(String color, XSSFSheet sheet) {
        if (StringUtils.isBlank(color)) {
            return;
        }
        String rgbHex = color.replace("#", "");
        byte[] rgbBytes = new byte[3];
        rgbBytes[0] = (byte) Integer.parseInt(rgbHex.substring(0, 2), 16); // R
        rgbBytes[1] = (byte) Integer.parseInt(rgbHex.substring(2, 4), 16); // G
        rgbBytes[2] = (byte) Integer.parseInt(rgbHex.substring(4, 6), 16); // B
        XSSFColor xssfColor = new XSSFColor(rgbBytes);
        // 设置标签颜色
        sheet.setTabColor(xssfColor);
    }

    private static void mapFrozen(Frozen frozen, XSSFSheet sheet) {
        if (frozen == null) {
            return;
        }

        switch (frozen.getType()) {
            case ROW:
                sheet.createFreezePane(0, 1);
                break;
            case COLUMN:
                sheet.createFreezePane(1, 0);
                break;
            case BOTH:
                sheet.createFreezePane(1, 1);
                break;
            case RANGE_ROW:
                sheet.createFreezePane(0, frozen.getRange().getRow_focus());
                break;
            case RANGE_COLUMN:
                sheet.createFreezePane(frozen.getRange().getColumn_focus(), 0);
                break;
            case RANGE_BOTH:
                sheet.createFreezePane(frozen.getRange().getColumn_focus(), frozen.getRange().getRow_focus());
                break;
            case CANCEL:
                sheet.createFreezePane(0, 0);
                break;
            default:
                break;
        }
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

    private static void mapDataVerification(Map<String, DataVerification> dataVerifications, XSSFSheet sheet) {
        // 注意：Luckysheet的数据验证是以"行_列"格式的键存储的，例如"1_0"表示第1行第0列
        // 这意味着每个验证都对应特定的单元格

        if (dataVerifications == null || dataVerifications.isEmpty()) {
            return;
        }

        // 遍历每个数据验证配置
        for (Map.Entry<String, DataVerification> entry : dataVerifications.entrySet()) {
            String key = entry.getKey();
            DataVerification dataVerification = entry.getValue();

            // 解析键以获取行列信息（格式为"row_col"）
            String[] parts = key.split("_");
            if (parts.length != 2) {
                continue;
            }

            try {
                int row = Integer.parseInt(parts[0]);
                int col = Integer.parseInt(parts[1]);

                org.apache.poi.ss.usermodel.DataValidationHelper validationHelper = sheet.getDataValidationHelper();

                String cellRef = toCellRef(row, col);
                org.apache.poi.ss.usermodel.DataValidationConstraint constraint = createConstraintFromLuckySheet(dataVerification, validationHelper, cellRef);

                if (constraint != null) {
                    // 创建指向特定单元格的数据验证
                    org.apache.poi.ss.util.CellRangeAddressList addressList = new org.apache.poi.ss.util.CellRangeAddressList(
                        row, row, col, col
                    );

                    // 创建POI数据验证对象
                    org.apache.poi.ss.usermodel.DataValidation dataValidation = validationHelper.createValidation(
                        constraint,
                        addressList
                    );

                    // 设置数据验证的其他属性
                    dataValidation.setEmptyCellAllowed(!dataVerification.getProhibitInput());
                    if (dataVerification.getHintText() != null && !dataVerification.getHintText().isEmpty()) {
                        dataValidation.createErrorBox("输入错误", dataVerification.getHintText());
                        dataValidation.createPromptBox("提示", dataVerification.getHintText());
                    }

                    // 将数据验证添加到工作表
                    sheet.addValidationData(dataValidation);
                }
            } catch (NumberFormatException e) {
                System.err.println("Invalid data verification key: " + key);
            } catch (Exception e) {
                System.err.println("Error processing data verification: " + key);
            }
        }
    }

    /**
     * 根据Luckysheet的DataVerification创建POI的DataValidationConstraint
     */
    private static org.apache.poi.ss.usermodel.DataValidationConstraint createConstraintFromLuckySheet(
            DataVerification dataVerification,
            org.apache.poi.ss.usermodel.DataValidationHelper validationHelper,
            String cellRef) {

        String type = dataVerification.getType();
        String type2 = dataVerification.getType2();
        Object value1 = dataVerification.getValue1();
        Object value2 = dataVerification.getValue2();

        try {
            switch (type) {
                case "dropdown":
                    // 下拉列表验证
                    if (value1 instanceof String) {
                        String[] values = ((String) value1).split(",");
                        return validationHelper.createExplicitListConstraint(values);
                    } else if (value1 instanceof String[]) {
                        return validationHelper.createExplicitListConstraint((String[]) value1);
                    }
                    break;

                case "number_integer":
                    return createIntConstraint(validationHelper, type2, value1, value2, cellRef);

                case "number_decimal":
                    return createDecimalConstraint(validationHelper, type2, value1, value2, cellRef);

                case "text_length":
                    return createTextLengthConstraint(validationHelper, type2, value1, value2, cellRef);

                case "date":
                    return createDateConstraint(validationHelper, type2, value1, value2, cellRef);

                case "number":
                    return createDecimalConstraint(validationHelper, type2, value1, value2, cellRef);

                default:
                    // 其他类型，默认使用自定义公式验证
                    if (value1 instanceof String) {
                        return validationHelper.createFormulaListConstraint((String) value1);
                    }
                    break;
            }
        } catch (Exception e) {
            System.err.println("Failed to create constraint: " + e.getMessage());
        }

        return null;
    }

    /**
     * 创建整数约束
     */
    private static org.apache.poi.ss.usermodel.DataValidationConstraint createIntConstraint(
            org.apache.poi.ss.usermodel.DataValidationHelper validationHelper,
            String operator, Object value1, Object value2, String cellRef) {

        int intValue1 = value1 != null ? convertToInt(value1) : 0;
        int intValue2 = value2 != null ? convertToInt(value2) : 0;

        switch (operator) {
            case "bw": // between
                return validationHelper.createIntegerConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    String.valueOf(intValue1), String.valueOf(intValue2));
            case "nb": // not between
                return validationHelper.createIntegerConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_BETWEEN,
                    String.valueOf(intValue1), String.valueOf(intValue2));
            case "eq": // equal
                return validationHelper.createIntegerConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.EQUAL,
                    String.valueOf(intValue1), null);
            case "ne": // not equal
                return validationHelper.createIntegerConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_EQUAL,
                    String.valueOf(intValue1), null);
            case "gt": // greater than
                return validationHelper.createIntegerConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.GREATER_THAN,
                    String.valueOf(intValue1), null);
            case "lt": // less than
                return validationHelper.createIntegerConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.LESS_THAN,
                    String.valueOf(intValue1), null);
            case "gte":
                return validationHelper.createCustomConstraint(cellRef + ">=" + intValue1);
            case "lte":
                return validationHelper.createCustomConstraint(cellRef + "<=" + intValue1);
            default:
                return validationHelper.createIntegerConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    String.valueOf(intValue1), String.valueOf(intValue2));
        }
    }

    /**
     * 创建小数约束
     */
    private static org.apache.poi.ss.usermodel.DataValidationConstraint createDecimalConstraint(
            org.apache.poi.ss.usermodel.DataValidationHelper validationHelper,
            String operator, Object value1, Object value2, String cellRef) {

        double decimalValue1 = value1 != null ? convertToDouble(value1) : 0.0;
        double decimalValue2 = value2 != null ? convertToDouble(value2) : 0.0;

        switch (operator) {
            case "bw": // between
                return validationHelper.createDecimalConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    String.valueOf(decimalValue1), String.valueOf(decimalValue2));
            case "nb": // not between
                return validationHelper.createDecimalConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_BETWEEN,
                    String.valueOf(decimalValue1), String.valueOf(decimalValue2));
            case "eq": // equal
                return validationHelper.createDecimalConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.EQUAL,
                    String.valueOf(decimalValue1), null);
            case "ne": // not equal
                return validationHelper.createDecimalConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_EQUAL,
                    String.valueOf(decimalValue1), null);
            case "gt": // greater than
                return validationHelper.createDecimalConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.GREATER_THAN,
                    String.valueOf(decimalValue1), null);
            case "lt": // less than
                return validationHelper.createDecimalConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.LESS_THAN,
                    String.valueOf(decimalValue1), null);
            case "gte":
                return validationHelper.createCustomConstraint(cellRef + ">=" + decimalValue1);
            case "lte":
                return validationHelper.createCustomConstraint(cellRef + "<=" + decimalValue1);
            default:
                return validationHelper.createDecimalConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    String.valueOf(decimalValue1), String.valueOf(decimalValue2));
        }
    }

    /**
     * 创建文本长度约束
     */
    private static org.apache.poi.ss.usermodel.DataValidationConstraint createTextLengthConstraint(
            org.apache.poi.ss.usermodel.DataValidationHelper validationHelper,
            String operator, Object value1, Object value2, String cellRef) {

        int lengthValue1 = value1 != null ? convertToInt(value1) : 0;
        int lengthValue2 = value2 != null ? convertToInt(value2) : 0;

        switch (operator) {
            case "bw": // between
                return validationHelper.createTextLengthConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    String.valueOf(lengthValue1), String.valueOf(lengthValue2));
            case "nb": // not between
                return validationHelper.createTextLengthConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_BETWEEN,
                    String.valueOf(lengthValue1), String.valueOf(lengthValue2));
            case "eq": // equal
                return validationHelper.createTextLengthConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.EQUAL,
                    String.valueOf(lengthValue1), null);
            case "ne": // not equal
                return validationHelper.createTextLengthConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_EQUAL,
                    String.valueOf(lengthValue1), null);
            case "gt": // greater than
                return validationHelper.createTextLengthConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.GREATER_THAN,
                    String.valueOf(lengthValue1), null);
            case "lt": // less than
                return validationHelper.createTextLengthConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.LESS_THAN,
                    String.valueOf(lengthValue1), null);
            case "gte":
                return validationHelper.createCustomConstraint(cellRef + ">=" + lengthValue1);
            case "lte":
                return validationHelper.createCustomConstraint(cellRef + "<=" + lengthValue1);
            default:
                return validationHelper.createTextLengthConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    String.valueOf(lengthValue1), String.valueOf(lengthValue2));
        }
    }

    /**
     * 创建日期约束
     */
    private static org.apache.poi.ss.usermodel.DataValidationConstraint createDateConstraint(
            org.apache.poi.ss.usermodel.DataValidationHelper validationHelper,
            String operator, Object value1, Object value2, String cellRef) {

        // 日期值通常以字符串形式表示，例如 "2023-01-01"
        String dateValue1 = value1 != null ? value1.toString() : "";
        String dateValue2 = value2 != null ? value2.toString() : "";

        switch (operator) {
            case "bw": // between
                return validationHelper.createDateConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    dateValue1, dateValue2, "yyyy-MM-dd");
            case "nb": // not between
                return validationHelper.createDateConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_BETWEEN,
                    dateValue1, dateValue2, "yyyy-MM-dd");
            case "eq": // equal
                return validationHelper.createDateConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.EQUAL,
                    dateValue1, null, "yyyy-MM-dd");
            case "ne": // not equal
                return validationHelper.createDateConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.NOT_EQUAL,
                    dateValue1, null, "yyyy-MM-dd");
            case "gt": // greater than
                return validationHelper.createDateConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.GREATER_THAN,
                    dateValue1, null, "yyyy-MM-dd");
            case "lt": // less than
                return validationHelper.createDateConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.LESS_THAN,
                    dateValue1, null, "yyyy-MM-dd");
            case "gte":
                return validationHelper.createCustomConstraint(cellRef + ">=" + dateValue1);
            case "lte":
                return validationHelper.createCustomConstraint(cellRef + "<=" + dateValue1);
            default:
                return validationHelper.createDateConstraint(
                    org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                    dateValue1, dateValue2, "yyyy-MM-dd");
        }
    }

    /**
     * 将对象转换为整数
     */
    private static int convertToInt(Object obj) {
        if (obj instanceof Number) {
            return ((Number) obj).intValue();
        } else if (obj instanceof String) {
            try {
                return Integer.parseInt((String) obj);
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        return 0;
    }

    /**
     * 将对象转换为双精度浮点数
     */
    private static double convertToDouble(Object obj) {
        if (obj instanceof Number) {
            return ((Number) obj).doubleValue();
        } else if (obj instanceof String) {
            try {
                return Double.parseDouble((String) obj);
            } catch (NumberFormatException e) {
                return 0.0;
            }
        }
        return 0.0;
    }

    private static String toCellRef(int row, int col) {
        StringBuilder colPart = new StringBuilder();
        int c = col;
        while (c >= 0) {
            colPart.insert(0, (char) ('A' + c % 26));
            c = c / 26 - 1;
        }
        return colPart.toString() + (row + 1);
    }
}
