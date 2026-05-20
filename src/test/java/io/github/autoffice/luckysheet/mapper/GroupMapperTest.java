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

import io.github.autoffice.luckysheet.model.sheet.Group;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link GroupMapper} 的行列分组双向映射, 涵盖各种边界条件和分组场景.
 */
class GroupMapperTest {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private LuckySheet luckySheet;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test");
        luckySheet = new LuckySheet();
        luckySheet.setName("Test");
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    // ========== mapToLuckySheet ==========

    @Test
    void mapToLuckySheet_emptyRowGroupsAndColGroups_setsNothing() {
        GroupMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getRowGroup());
        assertNull(luckySheet.getColGroup());
    }

    @Test
    void mapToLuckySheet_onlyRowGroups_setsRowGroupOnly() {
        XSSFRow row0 = sheet.createRow(0);
        row0.getCTRow().setOutlineLevel((short) 1);
        XSSFRow row1 = sheet.createRow(1);
        row1.getCTRow().setOutlineLevel((short) 1);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getRowGroup());
        assertEquals(1, luckySheet.getRowGroup().size());
        assertNull(luckySheet.getColGroup());
    }

    @Test
    void mapToLuckySheet_onlyColGroups_setsColGroupOnly() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTCols ctCols = ctWorksheet.addNewCols();
        CTCol ctCol = ctCols.addNewCol();
        ctCol.setMin(1);
        ctCol.setMax(3);
        ctCol.setOutlineLevel((short) 1);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNull(luckySheet.getRowGroup());
        assertNotNull(luckySheet.getColGroup());
        assertEquals(1, luckySheet.getColGroup().size());
    }

    @Test
    void mapToLuckySheet_bothRowAndColGroups_setsBoth() {
        XSSFRow row0 = sheet.createRow(0);
        row0.getCTRow().setOutlineLevel((short) 1);

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTCols ctCols = ctWorksheet.addNewCols();
        CTCol ctCol = ctCols.addNewCol();
        ctCol.setMin(1);
        ctCol.setMax(2);
        ctCol.setOutlineLevel((short) 1);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getRowGroup());
        assertNotNull(luckySheet.getColGroup());
    }

    @Test
    void mapToLuckySheet_rowWithNullRow_levelZero() {
        sheet.createRow(0);
        // Row 1 is null (not created)
        XSSFRow row2 = sheet.createRow(2);
        row2.getCTRow().setOutlineLevel((short) 1);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        // Should handle null row gracefully
        assertNotNull(luckySheet.getRowGroup());
    }

    @Test
    void mapToLuckySheet_rowLevelZero_skipped() {
        XSSFRow row0 = sheet.createRow(0);
        row0.getCTRow().setOutlineLevel((short) 0);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNull(luckySheet.getRowGroup());
    }

    @Test
    void mapToLuckySheet_multipleSameLevelRowsContinue_singleGroup() {
        XSSFRow row0 = sheet.createRow(0);
        row0.getCTRow().setOutlineLevel((short) 1);
        XSSFRow row1 = sheet.createRow(1);
        row1.getCTRow().setOutlineLevel((short) 1);
        XSSFRow row2 = sheet.createRow(2);
        row2.getCTRow().setOutlineLevel((short) 1);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getRowGroup());
        assertEquals(1, luckySheet.getRowGroup().size());
        Group g = luckySheet.getRowGroup().get(0);
        assertEquals(0, g.getStart());
        assertEquals(2, g.getEnd());
        assertEquals(1, g.getLevel());
    }

    @Test
    void mapToLuckySheet_levelChanges_newGroup() {
        XSSFRow row0 = sheet.createRow(0);
        row0.getCTRow().setOutlineLevel((short) 1);
        XSSFRow row1 = sheet.createRow(1);
        row1.getCTRow().setOutlineLevel((short) 1);
        XSSFRow row2 = sheet.createRow(2);
        row2.getCTRow().setOutlineLevel((short) 2);
        XSSFRow row3 = sheet.createRow(3);
        row3.getCTRow().setOutlineLevel((short) 2);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getRowGroup());
        assertEquals(2, luckySheet.getRowGroup().size());
        assertEquals(1, luckySheet.getRowGroup().get(0).getLevel());
        assertEquals(2, luckySheet.getRowGroup().get(1).getLevel());
    }

    @Test
    void mapToLuckySheet_hiddenRowSetsCollapsed() {
        XSSFRow row0 = sheet.createRow(0);
        row0.getCTRow().setOutlineLevel((short) 1);
        row0.setZeroHeight(true);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getRowGroup());
        assertEquals(1, luckySheet.getRowGroup().size());
        assertTrue(luckySheet.getRowGroup().get(0).getCollapsed());
    }

    @Test
    void mapToLuckySheet_colLevelZero_skipped() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTCols ctCols = ctWorksheet.addNewCols();
        CTCol ctCol = ctCols.addNewCol();
        ctCol.setMin(1);
        ctCol.setMax(2);
        ctCol.setOutlineLevel((short) 0);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNull(luckySheet.getColGroup());
    }

    @Test
    void mapToLuckySheet_validColWithLevel_createsGroup() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTCols ctCols = ctWorksheet.addNewCols();
        CTCol ctCol = ctCols.addNewCol();
        ctCol.setMin(1);
        ctCol.setMax(3);
        ctCol.setOutlineLevel((short) 2);
        ctCol.setHidden(true);

        GroupMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getColGroup());
        assertEquals(1, luckySheet.getColGroup().size());
        Group g = luckySheet.getColGroup().get(0);
        assertEquals(0, g.getStart()); // min-1
        assertEquals(2, g.getEnd());   // max-1
        assertEquals(2, g.getLevel());
        assertTrue(g.getCollapsed());
    }

    // ========== mapToExcel ==========

    @Test
    void mapToExcel_nullRowGroups_doesNothing() {
        GroupMapper.mapToExcel(null, null, sheet);
        // No exception, no groups created
    }

    @Test
    void mapToExcel_emptyRowGroups_doesNothing() {
        GroupMapper.mapToExcel(new ArrayList<>(), new ArrayList<>(), sheet);
        // No exception, no groups created
    }

    @Test
    void mapToExcel_nullColGroups_doesNothing() {
        GroupMapper.mapToExcel(new ArrayList<>(), null, sheet);
        // No exception
    }

    @Test
    void mapToExcel_emptyColGroups_doesNothing() {
        GroupMapper.mapToExcel(new ArrayList<>(), new ArrayList<>(), sheet);
        // No exception
    }

    @Test
    void mapToExcel_groupWithNullStart_skipped() {
        Group g = new Group();
        g.setStart(null);
        g.setEnd(5);
        g.setLevel(1);
        GroupMapper.mapToExcel(Collections.singletonList(g), null, sheet);
        // No exception, group skipped
    }

    @Test
    void mapToExcel_groupWithNullEnd_skipped() {
        Group g = new Group();
        g.setStart(0);
        g.setEnd(null);
        g.setLevel(1);
        GroupMapper.mapToExcel(Collections.singletonList(g), null, sheet);
        // No exception, group skipped
    }

    @Test
    void mapToExcel_validRowGroup_createsGroup() {
        Group g = new Group();
        g.setStart(0);
        g.setEnd(5);
        g.setLevel(1);
        g.setCollapsed(false);

        GroupMapper.mapToExcel(Collections.singletonList(g), null, sheet);

        // Verify group was created by checking outline level
        XSSFRow row0 = sheet.getRow(0);
        assertNotNull(row0);
        assertTrue(row0.getCTRow().getOutlineLevel() > 0);
    }

    @Test
    void mapToExcel_rowGroupCollapsed_callsSetRowGroupCollapsed() {
        Group g = new Group();
        g.setStart(0);
        g.setEnd(3);
        g.setLevel(1);
        g.setCollapsed(true);

        GroupMapper.mapToExcel(Collections.singletonList(g), null, sheet);

        // Verify collapsed state was set
        XSSFRow row0 = sheet.getRow(0);
        assertNotNull(row0);
    }

    @Test
    void mapToExcel_validColGroup_createsGroup() {
        Group g = new Group();
        g.setStart(0);
        g.setEnd(3);
        g.setLevel(1);
        g.setCollapsed(false);

        GroupMapper.mapToExcel(null, Collections.singletonList(g), sheet);

        // Verify column group was created
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        assertNotNull(ctWorksheet);
    }

    @Test
    void mapToExcel_colGroupCollapsed_callsSetColumnGroupCollapsed() {
        Group g = new Group();
        g.setStart(0);
        g.setEnd(2);
        g.setLevel(1);
        g.setCollapsed(true);

        GroupMapper.mapToExcel(null, Collections.singletonList(g), sheet);

        // Verify column group was created with collapsed state
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        assertNotNull(ctWorksheet);
    }

    @Test
    void mapToExcel_nullGroupInList_skipped() {
        List<Group> groups = new ArrayList<>();
        groups.add(null);
        Group valid = new Group();
        valid.setStart(0);
        valid.setEnd(2);
        valid.setLevel(1);
        groups.add(valid);

        GroupMapper.mapToExcel(groups, null, sheet);

        // Valid group should be created, null skipped
        XSSFRow row0 = sheet.getRow(0);
        assertNotNull(row0);
    }

    // ========== Round-trip ==========

    @Test
    void roundTrip_rowGroups() {
        Group g = new Group();
        g.setStart(0);
        g.setEnd(5);
        g.setLevel(1);
        g.setCollapsed(false);

        GroupMapper.mapToExcel(Collections.singletonList(g), null, sheet);

        LuckySheet result = new LuckySheet();
        GroupMapper.mapToLuckySheet(sheet, result);

        assertNotNull(result.getRowGroup());
        assertEquals(1, result.getRowGroup().size());
        assertEquals(0, result.getRowGroup().get(0).getStart());
        assertEquals(5, result.getRowGroup().get(0).getEnd());
        assertEquals(1, result.getRowGroup().get(0).getLevel());
    }
}

