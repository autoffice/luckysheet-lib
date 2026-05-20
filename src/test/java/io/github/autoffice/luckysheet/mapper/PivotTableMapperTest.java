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
import io.github.autoffice.luckysheet.model.sheet.Range;
import io.github.autoffice.luckysheet.model.sheet.SheetPivotTable;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link PivotTableMapper} 的双向映射逻辑, 覆盖各种边界条件和数据透视表配置.
 */
class PivotTableMapperTest {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private XSSFSheet dataSheet;
    private LuckySheet luckySheet;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Pivot");
        dataSheet = workbook.createSheet("Data");
        luckySheet = new LuckySheet();
        luckySheet.setName("Pivot");
        luckySheet.setIndex("0");

        // 创建示例数据用于透视表
        createSampleData(dataSheet);
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    private void createSampleData(XSSFSheet sheet) {
        // 表头
        sheet.createRow(0).createCell(0).setCellValue("Category");
        sheet.getRow(0).createCell(1).setCellValue("Product");
        sheet.getRow(0).createCell(2).setCellValue("Sales");

        // 数据行
        sheet.createRow(1).createCell(0).setCellValue("A");
        sheet.getRow(1).createCell(1).setCellValue("P1");
        sheet.getRow(1).createCell(2).setCellValue(100);

        sheet.createRow(2).createCell(0).setCellValue("B");
        sheet.getRow(2).createCell(1).setCellValue("P2");
        sheet.getRow(2).createCell(2).setCellValue(200);
    }

    // ========== mapToLuckySheet tests ==========

    @Test
    void mapToLuckySheet_noPivotTables_doesNothing() {
        PivotTableMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getPivotTable());
    }

    @Test
    void mapToLuckySheet_emptyPivotTableList_doesNothing() {
        // 空列表情况已由 POI 内部处理，此处验证无透视表时的行为
        PivotTableMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getPivotTable());
    }

    @Test
    void mapToLuckySheet_withOnePivotTable() throws Exception {
        // 在数据表上创建透视表
        AreaReference source = new AreaReference("A1:C3",
                workbook.getSpreadsheetVersion());
        CellReference position = new CellReference(0, 0);
        XSSFPivotTable pivotTable = sheet.createPivotTable(source, position, dataSheet);

        PivotTableMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getPivotTable());
        assertTrue(luckySheet.getIsPivotTable());
        assertNotNull(luckySheet.getPivotTable().getPivot_select_save());
    }

    @Test
    void mapToLuckySheet_multiplePivotTables_takesFirst() throws Exception {
        // 创建两个透视表
        AreaReference source1 = new AreaReference("A1:C3",
                workbook.getSpreadsheetVersion());
        sheet.createPivotTable(source1, new CellReference(0, 0), dataSheet);

        AreaReference source2 = new AreaReference("A1:C3",
                workbook.getSpreadsheetVersion());
        sheet.createPivotTable(source2, new CellReference(10, 0), dataSheet);

        PivotTableMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getPivotTable());
        // 只取第一个透视表
        assertTrue(luckySheet.getIsPivotTable());
    }

    // ========== mapToExcel tests ==========

    @Test
    void mapToExcel_nullPivotTable_doesNothing() {
        PivotTableMapper.mapToExcel(null, sheet);
        assertEquals(0, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_invalidSheetIndex_negative() {
        SheetPivotTable pivotTable = new SheetPivotTable();
        pivotTable.setPivotDataSheetIndex(-1);
        pivotTable.setPivot_select_save(buildRange(0, 2, 0, 2));

        PivotTableMapper.mapToExcel(pivotTable, sheet);
        assertEquals(0, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_invalidSheetIndex_outOfRange() {
        SheetPivotTable pivotTable = new SheetPivotTable();
        pivotTable.setPivotDataSheetIndex(999);
        pivotTable.setPivot_select_save(buildRange(0, 2, 0, 2));

        PivotTableMapper.mapToExcel(pivotTable, sheet);
        assertEquals(0, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_nullPivotSelectSave() {
        SheetPivotTable pivotTable = new SheetPivotTable();
        pivotTable.setPivotDataSheetIndex(1);
        pivotTable.setPivot_select_save(null);

        PivotTableMapper.mapToExcel(pivotTable, sheet);
        assertEquals(0, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_incompletePivotSelectSave_rowLessThan2() {
        SheetPivotTable pivotTable = new SheetPivotTable();
        pivotTable.setPivotDataSheetIndex(1);
        Range range = new Range();
        range.setRow(Collections.singletonList(0));
        range.setColumn(Arrays.asList(0, 2));
        pivotTable.setPivot_select_save(range);

        PivotTableMapper.mapToExcel(pivotTable, sheet);
        assertEquals(0, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_incompletePivotSelectSave_columnLessThan2() {
        SheetPivotTable pivotTable = new SheetPivotTable();
        pivotTable.setPivotDataSheetIndex(1);
        Range range = new Range();
        range.setRow(Arrays.asList(0, 2));
        range.setColumn(Collections.singletonList(0));
        pivotTable.setPivot_select_save(range);

        PivotTableMapper.mapToExcel(pivotTable, sheet);
        assertEquals(0, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_withRowLabels() {
        SheetPivotTable pivotTable = createBasicPivotTable();
        List<SheetPivotTable.PivotColRow> rows = new ArrayList<>();
        SheetPivotTable.PivotColRow row = new SheetPivotTable.PivotColRow();
        row.setIndex(0);
        row.setName("Category");
        rows.add(row);
        pivotTable.setRow(rows);

        PivotTableMapper.mapToExcel(pivotTable, sheet);
        assertEquals(1, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_withColumnLabels() {
        SheetPivotTable pivotTable = createBasicPivotTable();
        List<SheetPivotTable.PivotColRow> cols = new ArrayList<>();
        SheetPivotTable.PivotColRow col = new SheetPivotTable.PivotColRow();
        col.setIndex(1);
        col.setName("Product");
        cols.add(col);
        pivotTable.setColumn(cols);

        PivotTableMapper.mapToExcel(pivotTable, sheet);
        assertEquals(1, sheet.getPivotTables().size());
    }

    @Test
    void mapToExcel_allSumtypes() {
        testSumtype("SUM");
        testSumtype("COUNT");
        testSumtype("COUNTA");
        testSumtype("AVG");
        testSumtype("AVERAGE");
        testSumtype("MAX");
        testSumtype("MIN");
        testSumtype("PRODUCT");
        testSumtype("STDEV");
        testSumtype("VAR");
        testSumtype("UNKNOWN");
        testSumtype(null);
    }

    private void testSumtype(String sumtype) {
        XSSFSheet testSheet = workbook.createSheet("Test_" + sumtype);
        SheetPivotTable pivotTable = new SheetPivotTable();
        pivotTable.setPivotDataSheetIndex(1);
        pivotTable.setPivot_select_save(buildRange(0, 2, 0, 2));

        SheetPivotTable.PivotValue val = new SheetPivotTable.PivotValue();
        val.setIndex(2);
        val.setName("Sales");
        val.setSumtype(sumtype);
        pivotTable.setValues(Collections.singletonList(val));

        PivotTableMapper.mapToExcel(pivotTable, testSheet);
        assertEquals(1, testSheet.getPivotTables().size());
    }

    // ========== Helper methods ==========

    private SheetPivotTable createBasicPivotTable() {
        SheetPivotTable pivotTable = new SheetPivotTable();
        pivotTable.setPivotDataSheetIndex(1);
        pivotTable.setPivot_select_save(buildRange(0, 2, 0, 2));
        return pivotTable;
    }

    private Range buildRange(int r1, int r2, int c1, int c2) {
        Range r = new Range();
        r.setRow(Arrays.asList(r1, r2));
        r.setColumn(Arrays.asList(c1, c2));
        return r;
    }
}
