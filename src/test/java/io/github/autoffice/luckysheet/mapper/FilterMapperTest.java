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

import io.github.autoffice.luckysheet.model.sheet.FilterColumn;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link FilterMapper} 的双向映射逻辑.
 */
class FilterMapperTest {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private LuckySheet luckySheet;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test");
        luckySheet = new LuckySheet();
        luckySheet.setName("Test");
        luckySheet.setIndex("0");
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    // ========== mapToExcel ==========

    @Test
    void mapToExcel_nullFilterSelect_doesNothing() {
        FilterMapper.mapToExcel(null, null, sheet);
        assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
    }

    @Test
    void mapToExcel_emptyRow_doesNothing() {
        Range r = new Range();
        r.setRow(Collections.emptyList());
        r.setColumn(Arrays.asList(0, 4));
        FilterMapper.mapToExcel(r, null, sheet);
        assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
    }

    @Test
    void mapToExcel_emptyColumn_doesNothing() {
        Range r = new Range();
        r.setRow(Arrays.asList(0, 9));
        r.setColumn(Collections.emptyList());
        FilterMapper.mapToExcel(r, null, sheet);
        assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
    }

    @Test
    void mapToExcel_singleElementRow_doesNothing() {
        Range r = new Range();
        r.setRow(Collections.singletonList(0));
        r.setColumn(Arrays.asList(0, 4));
        FilterMapper.mapToExcel(r, null, sheet);
        assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
    }

    @Test
    void mapToExcel_singleElementColumn_doesNothing() {
        Range r = new Range();
        r.setRow(Arrays.asList(0, 9));
        r.setColumn(Collections.singletonList(0));
        FilterMapper.mapToExcel(r, null, sheet);
        assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
    }

    @Test
    void mapToExcel_validRange_setsAutoFilter() {
        Range r = new Range();
        r.setRow(Arrays.asList(0, 9));
        r.setColumn(Arrays.asList(0, 4));
        FilterMapper.mapToExcel(r, null, sheet);
        assertTrue(sheet.getCTWorksheet().isSetAutoFilter());
        assertEquals("A1:E10", sheet.getCTWorksheet().getAutoFilter().getRef());
    }

    @Test
    void mapToExcel_withFilterColumns_stillSetsRange() {
        Range r = new Range();
        r.setRow(Arrays.asList(0, 9));
        r.setColumn(Arrays.asList(0, 4));
        Map<String, FilterColumn> filters = new HashMap<>();
        FilterColumn fc = new FilterColumn();
        fc.setStr("textFilter");
        fc.setStringsArray(Arrays.asList("a", "b"));
        filters.put("0", fc);
        FilterMapper.mapToExcel(r, filters, sheet);
        // Range is set; per implementation note, column conditions are not written by design.
        assertTrue(sheet.getCTWorksheet().isSetAutoFilter());
    }

    // ========== mapToLuckySheet ==========

    @Test
    void mapToLuckySheet_noAutoFilter_doesNothing() {
        FilterMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getFilter_select());
    }

    @Test
    void mapToLuckySheet_withAutoFilter_setsRange() {
        sheet.setAutoFilter(new CellRangeAddress(0, 9, 0, 4));
        FilterMapper.mapToLuckySheet(sheet, luckySheet);

        Range fs = luckySheet.getFilter_select();
        assertNotNull(fs);
        assertEquals(Arrays.asList(0, 9), fs.getRow());
        assertEquals(Arrays.asList(0, 4), fs.getColumn());
    }

    @Test
    void mapToLuckySheet_singleRowRange() {
        sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, 4));
        FilterMapper.mapToLuckySheet(sheet, luckySheet);

        Range fs = luckySheet.getFilter_select();
        assertNotNull(fs);
        assertEquals(Arrays.asList(0, 0), fs.getRow());
    }

    // ========== Round-trip ==========

    @Test
    void roundTrip_filterSelectRange() {
        Range src = new Range();
        src.setRow(Arrays.asList(2, 7));
        src.setColumn(Arrays.asList(1, 3));
        FilterMapper.mapToExcel(src, null, sheet);

        LuckySheet result = new LuckySheet();
        FilterMapper.mapToLuckySheet(sheet, result);

        Range fs = result.getFilter_select();
        assertNotNull(fs);
        assertEquals(Arrays.asList(2, 7), fs.getRow());
        assertEquals(Arrays.asList(1, 3), fs.getColumn());
    }
}
