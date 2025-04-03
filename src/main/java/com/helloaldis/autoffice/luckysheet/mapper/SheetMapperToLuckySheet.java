package com.helloaldis.autoffice.luckysheet.mapper;

import com.helloaldis.autoffice.luckysheet.model.cell.CellData;
import com.helloaldis.autoffice.luckysheet.model.cell.MergeCell;
import com.helloaldis.autoffice.luckysheet.model.sheet.BoolStatus;
import com.helloaldis.autoffice.luckysheet.model.sheet.Border;
import com.helloaldis.autoffice.luckysheet.model.sheet.LuckySheet;
import com.helloaldis.autoffice.luckysheet.util.NumberUtil;
import com.helloaldis.autoffice.luckysheet.util.PoiUtil;
import org.apache.poi.ss.util.CellRangeAddress;
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

        ImageMapperToLuckySheet.mapToSheet(sheet, luckySheet);
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
