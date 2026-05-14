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
import io.github.autoffice.luckysheet.model.sheet.Frozen;
import io.github.autoffice.luckysheet.model.sheet.FrozenType;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.util.NumberUtil;
import io.github.autoffice.luckysheet.util.PoiUtil;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class SheetMapperToLuckySheet {
    public static void mapToSheet(XSSFSheet sheet, LuckySheet luckySheet) {
        luckySheet.setZoomRatio(1.0);

        mapDefaultRowHeight(sheet, luckySheet);
        mapDefaultColumnWidth(sheet, luckySheet);

        for (int rowIndex = PoiUtil.getFirstRowBase0(sheet); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            XSSFRow row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            for (int colIndex = row.getFirstCellNum(); colIndex < row.getLastCellNum(); colIndex++) {
                XSSFCell cell = row.getCell(colIndex);
                if (cell == null) {
                    continue;
                }

                CellData cellData = LuckySheetFactory.createCellData();
                cellData.setC(colIndex);
                cellData.setR(rowIndex);
                CellMapperToLuckySheet.mapToCell(cell, cellData);
                luckySheet.getCelldata().add(cellData);

                mapBorder(cell, luckySheet);
            }
        }

        mapSheetStatus(sheet, luckySheet);
        mapSheetHidden(sheet, luckySheet);
        mapMergeCell(sheet, luckySheet);
        mapRowHeightAndHidden(sheet, luckySheet);
        mapColumnWithAndHidden(sheet, luckySheet);
        mapGridLines(sheet, luckySheet);
        mapFrozen(sheet, luckySheet);
        mapDataVerification(sheet, luckySheet);

        ImageMapperToLuckySheet.mapToSheet(sheet, luckySheet);
    }

    private static void mapDataVerification(XSSFSheet sheet, LuckySheet luckySheet) {
        // 获取工作表中的所有数据验证
        List<org.apache.poi.xssf.usermodel.XSSFDataValidation> xssfDataValidations = sheet.getDataValidations();
        if (xssfDataValidations == null || xssfDataValidations.isEmpty()) {
            return;
        }

        // 初始化luckysheet的数据验证映射
        if (luckySheet.getDataVerification() == null) {
            luckySheet.setDataVerification(new java.util.HashMap<>());
        }

        // 遍历每个数据验证规则
        for (org.apache.poi.xssf.usermodel.XSSFDataValidation dataValidation : xssfDataValidations) {
            DataValidationConstraint constraint = dataValidation.getValidationConstraint();
            if (constraint == null) {
                continue; // 跳过无效的约束
            }

            // 获取数据验证应用的单元格范围
            org.apache.poi.ss.util.CellRangeAddressList regions = dataValidation.getRegions();

            // 遍历每个应用验证的单元格范围
            for (int i = 0; i < regions.countRanges(); i++) {
                CellRangeAddress region = regions.getCellRangeAddress(i);

                // 对于每个单元格范围，创建对应的数据验证项
                for (int row = region.getFirstRow(); row <= region.getLastRow(); row++) {
                    for (int col = region.getFirstColumn(); col <= region.getLastColumn(); col++) {
                        // 创建Luckysheet数据验证对象
                        DataVerification verification = createDataVerification(dataValidation, constraint);
                        if (verification != null) {
                            // 使用 "row_col" 格式作为键
                            String key = row + "_" + col;
                            // 添加到luckysheet的数据验证映射
                            luckySheet.getDataVerification().put(key, verification);
                        }
                    }
                }
            }
        }
    }

    private static DataVerification createDataVerification(org.apache.poi.xssf.usermodel.XSSFDataValidation dataValidation, DataValidationConstraint constraint) {
        int validationType = constraint.getValidationType();
        if (validationType == DataValidationConstraint.ValidationType.ANY) {
            return null; // 任意值不需要验证
        }

        DataVerification verification = new DataVerification();

        switch (validationType) {
            case DataValidationConstraint.ValidationType.INTEGER:
                verification.setType("number_integer");
                break;
            case DataValidationConstraint.ValidationType.DECIMAL:
                verification.setType("number_decimal");
                break;
            case DataValidationConstraint.ValidationType.TEXT_LENGTH:
                verification.setType("text_length");
                break;
            case DataValidationConstraint.ValidationType.LIST:
                verification.setType("dropdown");
                break;
            case DataValidationConstraint.ValidationType.DATE:
                verification.setType("date");
                break;
            case DataValidationConstraint.ValidationType.TIME:
                verification.setType("time");
                break;
            case DataValidationConstraint.ValidationType.FORMULA:
            default:
                verification.setType("validity");
                break;
        }

        setConditionType(verification, constraint);

        Object[] explicitListValues = constraint.getExplicitListValues();
        if (explicitListValues != null && explicitListValues.length > 0) {
            // 根据Luckysheet规范，如果是下拉列表类型，需要将所有选项合并为逗号分隔的字符串
            if ("dropdown".equals(verification.getType())) {
                StringBuilder sb = new StringBuilder();
                for (Object explicitListValue : explicitListValues) {
                    if (explicitListValue != null) {
                        if (sb.length() > 0) {
                            sb.append(",");
                        }
                        sb.append(explicitListValue);
                    }
                }
                if (sb.length() > 0) {
                    verification.setValue1(sb.toString());
                }
            } else {
                // 对于非下拉列表类型，按照原有逻辑处理
                if (explicitListValues[0] != null) {
                    verification.setValue1(explicitListValues[0]);
                    // 根据Luckysheet规范，只有在type2为"bw"、"nb"或"type"为"checkbox"时才设置value2
                    if (explicitListValues.length >= 2 && explicitListValues[1] != null &&
                        ("bw".equals(verification.getType2()) || "nb".equals(verification.getType2()))) {
                        verification.setValue2(explicitListValues[1]);
                    }
                }
            }
        } else {
            String formula1 = constraint.getFormula1();
            String formula2 = constraint.getFormula2();
            if (formula1 != null) {
                verification.setValue1(formula1);
            }
            // 同样根据Luckysheet规范设置value2
            if (formula2 != null &&
                ("bw".equals(verification.getType2()) || "nb".equals(verification.getType2()))) {
                verification.setValue2(formula2);
            }
        }

        // 设置prohibitInput，默认为false
        boolean prohibitInput = !dataValidation.getEmptyCellAllowed();
        if (prohibitInput) {
            verification.setProhibitInput(prohibitInput);
        }

        // 设置hintShow和hintText
        boolean hintShow = dataValidation.getShowErrorBox();
        if (hintShow) {
            verification.setHintShow(hintShow);
            String errorTitle = dataValidation.getErrorBoxTitle();
            String errorText = dataValidation.getErrorBoxText();
            if (errorTitle != null && !errorTitle.isEmpty()) {
                verification.setHintText(errorTitle);
            } else if (errorText != null && !errorText.isEmpty()) {
                verification.setHintText(errorText);
            }
        }

        return verification;
    }

    private static void setConditionType(DataVerification verification, DataValidationConstraint constraint) {
        int validationType = constraint.getValidationType();
        switch (validationType) {
            case DataValidationConstraint.ValidationType.INTEGER:
            case DataValidationConstraint.ValidationType.DECIMAL:
            case DataValidationConstraint.ValidationType.TEXT_LENGTH:
                setOperatorConditionType(verification, constraint);
                break;
            case DataValidationConstraint.ValidationType.DATE:
                setDateConditionType(verification, constraint);
                break;
            case DataValidationConstraint.ValidationType.LIST:
                //verification.setType2("dropdown");
                break;
            default:
                verification.setType2("bw");
                break;
        }
    }

    private static void setDateConditionType(DataVerification verification, DataValidationConstraint constraint) {
        switch (constraint.getOperator()) {
            case DataValidationConstraint.OperatorType.BETWEEN:
                verification.setType2("bw");
                break;
            case DataValidationConstraint.OperatorType.NOT_BETWEEN:
                verification.setType2("nb");
                break;
            case DataValidationConstraint.OperatorType.EQUAL:
                verification.setType2("eq");
                break;
            case DataValidationConstraint.OperatorType.NOT_EQUAL:
                verification.setType2("ne");
                break;
            case DataValidationConstraint.OperatorType.GREATER_THAN:
                verification.setType2("af"); // 晚于
                break;
            case DataValidationConstraint.OperatorType.LESS_THAN:
                verification.setType2("bf"); // 早于
                break;

            default:
                verification.setType2("bw");
                break;
        }
    }

    private static void setOperatorConditionType(DataVerification verification, DataValidationConstraint constraint) {
        switch (constraint.getOperator()) {
            case DataValidationConstraint.OperatorType.BETWEEN:
                verification.setType2("bw");
                break;
            case DataValidationConstraint.OperatorType.NOT_BETWEEN:
                verification.setType2("nb");
                break;
            case DataValidationConstraint.OperatorType.EQUAL:
                verification.setType2("eq");
                break;
            case DataValidationConstraint.OperatorType.NOT_EQUAL:
                verification.setType2("ne");
                break;
            case DataValidationConstraint.OperatorType.GREATER_THAN:
                verification.setType2("gt");
                break;
            case DataValidationConstraint.OperatorType.LESS_THAN:
                verification.setType2("lt");
                break;

            default:
                verification.setType2("bw");
                break;
        }
    }

    private static void mapFrozen(XSSFSheet sheet, LuckySheet luckySheet) {
        PaneInformation paneInformation = sheet.getPaneInformation();
        if (paneInformation == null || !paneInformation.isFreezePane()) {
            return;
        }

        short topRow = paneInformation.getHorizontalSplitPosition();
        short leftCol = paneInformation.getVerticalSplitPosition();
        if (topRow == 0 && leftCol == 0) {
            return;
        }

        Frozen frozen = LuckySheetFactory.createFrozen();
        if (topRow == 1 && leftCol == 0) {
            frozen.setType(FrozenType.ROW);
        } else if (topRow == 0 && leftCol == 1) {
            frozen.setType(FrozenType.COLUMN);
        } else if (topRow == 1 && leftCol == 1) {
            frozen.setType(FrozenType.BOTH);
        } else if (topRow == 0 && leftCol > 1) {
            frozen.setType(FrozenType.RANGE_COLUMN);
            frozen.setRange(new Frozen.Range(topRow, leftCol));
        } else if (topRow > 1 && leftCol == 0) {
            frozen.setType(FrozenType.RANGE_ROW);
            frozen.setRange(new Frozen.Range(topRow, leftCol));
        } else if (topRow > 1 && leftCol > 1) {
            frozen.setType(FrozenType.RANGE_BOTH);
            frozen.setRange(new Frozen.Range(topRow, leftCol));
        }

        luckySheet.setFrozen(frozen);
    }

    private static void mapGridLines(XSSFSheet sheet, LuckySheet luckySheet) {
        luckySheet.setShowGridLines(BoolStatus.of(sheet.isDisplayGridlines()));
    }

    private static void mapBorder(XSSFCell cell, LuckySheet luckySheet) {
        XSSFCellStyle cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            return;
        }
        Border border = LuckySheetFactory.createBorder();
        border.getValue().setCol_index(cell.getColumnIndex());
        border.getValue().setRow_index(cell.getRowIndex());
        border.getValue().setT(LuckySheetFactory.createBorderStyle(cellStyle.getTopBorderXSSFColor(), cellStyle.getBorderTop()));
        border.getValue().setR(LuckySheetFactory.createBorderStyle(cellStyle.getRightBorderXSSFColor(), cellStyle.getBorderRight()));
        border.getValue().setB(LuckySheetFactory.createBorderStyle(cellStyle.getBottomBorderXSSFColor(), cellStyle.getBorderBottom()));
        border.getValue().setL(LuckySheetFactory.createBorderStyle(cellStyle.getLeftBorderXSSFColor(), cellStyle.getBorderLeft()));

        if (LuckySheetFactory.hasBorderStyle(border)) {
            luckySheet.getConfig().getBorderInfo().add(border);
        }
    }

    private static void mapColumnWithAndHidden(XSSFSheet sheet, LuckySheet luckySheet) {
        short maxColNum = PoiUtil.getMaxColNum(sheet);
        for (int i = 0; i < maxColNum; i++) {
            float columnWidthInPixels = sheet.getColumnWidthInPixels(i);
            luckySheet.getConfig().getColumnlen().put(i, Math.round(columnWidthInPixels));


            if (sheet.isColumnHidden(i)) {
                luckySheet.getConfig().getColhidden().put(i, 0);
            }
        }
    }

    private static void mapRowHeightAndHidden(XSSFSheet sheet, LuckySheet luckySheet) {
        for (int i = PoiUtil.getFirstRowBase0(sheet); i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                luckySheet.getConfig().getRowlen().put(i, NumberUtil.twips2Pixel(row.getHeight()));
                if (row.getZeroHeight()) {
                    luckySheet.getConfig().getRowhidden().put(i, 0);
                }
            }
        }
    }

    private static void mapMergeCell(XSSFSheet sheet, LuckySheet luckySheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress mergedRegion : mergedRegions) {
            MergeCell mergeCell = new MergeCell();
            mergeCell.setR(mergedRegion.getFirstRow());
            mergeCell.setRs(mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1);
            mergeCell.setC(mergedRegion.getFirstColumn());
            mergeCell.setCs(mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1);
            luckySheet.getConfig().getMerge().put(mergeCell.getR() + "_" + mergeCell.getC(), mergeCell);
        }
    }

    private static void mapSheetHidden(XSSFSheet sheet, LuckySheet luckySheet) {
        int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
        luckySheet.setHide(BoolStatus.of(sheet.getWorkbook().isSheetHidden(sheetIndex)));
    }

    private static void mapSheetStatus(XSSFSheet sheet, LuckySheet luckySheet) {
        luckySheet.setStatus(BoolStatus.of(sheet.isSelected()));
    }

    private static void mapDefaultColumnWidth(XSSFSheet sheet, LuckySheet luckySheet) {
        luckySheet.setDefaultColWidth((short) NumberUtil.characterLen2Pixel(sheet.getDefaultColumnWidth()));
    }

    private static void mapDefaultRowHeight(XSSFSheet sheet, LuckySheet luckySheet) {
        luckySheet.setDefaultRowHeight((short) NumberUtil.twips2Pixel(sheet.getDefaultRowHeight()));
    }
}
