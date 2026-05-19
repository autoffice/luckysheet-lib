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

import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.SheetPivotTable;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

/**
 * 数据透视表 Luckysheet ↔ POI 双向映射器 (有损转换).
 */
public final class PivotTableMapper {

    private PivotTableMapper() {
    }

    /**
     * 从 Excel 工作表中提取数据透视表信息并转换为 Luckysheet 格式.
     *
     * @param sheet      源 POI 工作表
     * @param luckySheet 目标 Luckysheet 工作表
     */
    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        List<XSSFPivotTable> tables = sheet.getPivotTables();
        if (tables == null || tables.isEmpty()) {
            return;
        }
        // Luckysheet LuckySheet.pivotTable 仅允许一个透视表, 取第一个
        XSSFPivotTable poiTable = tables.get(0);
        SheetPivotTable model = new SheetPivotTable();
        try {
            String ref = poiTable.getPivotCacheDefinition().getCTPivotCacheDefinition()
                    .getCacheSource().getWorksheetSource().getRef();
            if (ref != null) {
                AreaReference source = new AreaReference(ref,
                        sheet.getWorkbook().getSpreadsheetVersion());
                model.setPivot_select_save(rangeOf(source));
            }
        } catch (Exception ignore) {
            // 源引用读不到时跳过, 保留空 model
            model.setPivot_select_save(null);
        }
        luckySheet.setIsPivotTable(true);
        luckySheet.setPivotTable(model);
    }

    /**
     * 将 Luckysheet 数据透视表写入 Excel 工作表.
     *
     * @param pivotTable Luckysheet 数据透视表模型
     * @param sheet      目标 POI 工作表
     */
    public static void mapToExcel(SheetPivotTable pivotTable, XSSFSheet sheet) {
        if (pivotTable == null) {
            return;
        }
        XSSFWorkbook workbook = sheet.getWorkbook();
        int srcIdx = pivotTable.getPivotDataSheetIndex() == null ? 0 : pivotTable.getPivotDataSheetIndex();
        if (srcIdx < 0 || srcIdx >= workbook.getNumberOfSheets()) {
            return;
        }
        XSSFSheet srcSheet = workbook.getSheetAt(srcIdx);
        AreaReference srcRef;
        try {
            srcRef = buildAreaReference(srcSheet, pivotTable);
        } catch (Exception e) {
            return;
        }
        if (srcRef == null) {
            return;
        }

        CellReference position = new CellReference(0, 0);
        XSSFPivotTable poiTable;
        try {
            poiTable = sheet.createPivotTable(srcRef, position, srcSheet);
        } catch (Exception e) {
            return;
        }
        if (CollectionUtils.isNotEmpty(pivotTable.getRow())) {
            for (SheetPivotTable.PivotColRow row : pivotTable.getRow()) {
                if (row != null && row.getIndex() != null) {
                    poiTable.addRowLabel(row.getIndex());
                }
            }
        }
        if (CollectionUtils.isNotEmpty(pivotTable.getColumn())) {
            for (SheetPivotTable.PivotColRow col : pivotTable.getColumn()) {
                if (col != null && col.getIndex() != null) {
                    poiTable.addColLabel(col.getIndex());
                }
            }
        }
        if (CollectionUtils.isNotEmpty(pivotTable.getValues())) {
            for (SheetPivotTable.PivotValue val : pivotTable.getValues()) {
                if (val == null || val.getIndex() == null) {
                    continue;
                }
                DataConsolidateFunction func = toConsolidateFunction(val.getSumtype());
                poiTable.addColumnLabel(func, val.getIndex(),
                        val.getName() == null ? "" : val.getName());
            }
        }
    }

    private static AreaReference buildAreaReference(XSSFSheet srcSheet, SheetPivotTable pivotTable) {
        if (pivotTable.getPivot_select_save() == null
                || CollectionUtils.size(pivotTable.getPivot_select_save().getRow()) < 2
                || CollectionUtils.size(pivotTable.getPivot_select_save().getColumn()) < 2) {
            return null;
        }
        int firstRow = pivotTable.getPivot_select_save().getRow().get(0);
        int lastRow = pivotTable.getPivot_select_save().getRow().get(1);
        int firstCol = pivotTable.getPivot_select_save().getColumn().get(0);
        int lastCol = pivotTable.getPivot_select_save().getColumn().get(1);
        CellReference firstCell = new CellReference(srcSheet.getSheetName(), firstRow, firstCol, true, true);
        CellReference lastCell = new CellReference(srcSheet.getSheetName(), lastRow, lastCol, true, true);
        return new AreaReference(firstCell, lastCell, srcSheet.getWorkbook().getSpreadsheetVersion());
    }

    private static io.github.autoffice.luckysheet.model.sheet.Range rangeOf(AreaReference src) {
        io.github.autoffice.luckysheet.model.sheet.Range r = new io.github.autoffice.luckysheet.model.sheet.Range();
        CellReference first = src.getFirstCell();
        CellReference last = src.getLastCell();
        r.setRow(java.util.Arrays.asList(first.getRow(), last.getRow()));
        r.setColumn(java.util.Arrays.asList((int) first.getCol(), (int) last.getCol()));
        return r;
    }

    private static DataConsolidateFunction toConsolidateFunction(String sumtype) {
        if (sumtype == null) {
            return DataConsolidateFunction.SUM;
        }
        switch (sumtype.toUpperCase()) {
            case "SUM":
                return DataConsolidateFunction.SUM;
            case "COUNT":
            case "COUNTA":
                return DataConsolidateFunction.COUNT;
            case "AVG":
            case "AVERAGE":
                return DataConsolidateFunction.AVERAGE;
            case "MAX":
                return DataConsolidateFunction.MAX;
            case "MIN":
                return DataConsolidateFunction.MIN;
            case "PRODUCT":
                return DataConsolidateFunction.PRODUCT;
            case "STDEV":
                return DataConsolidateFunction.STD_DEV;
            case "VAR":
                return DataConsolidateFunction.VAR;
            default:
                return DataConsolidateFunction.SUM;
        }
    }
}
