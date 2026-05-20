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

import io.github.autoffice.luckysheet.model.sheet.ConditionFormat;
import io.github.autoffice.luckysheet.model.sheet.ConditionFormatType;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFPatternFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link ConditionFormatMapper} 的双向映射逻辑, 重点覆盖各种条件格式类型与
 * 比较运算符分支.
 */
class ConditionFormatMapperTest {

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

    // ========== mapToExcel: edge cases ==========

    @Test
    void mapToExcel_nullList_doesNothing() {
        ConditionFormatMapper.mapToExcel(null, sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    @Test
    void mapToExcel_emptyList_doesNothing() {
        ConditionFormatMapper.mapToExcel(new ArrayList<>(), sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    @Test
    void mapToExcel_nullCellrange_skipped() {
        ConditionFormat cf = new ConditionFormat();
        cf.setType(ConditionFormatType.DEFAULT);
        cf.setCellrange(null);
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    @Test
    void mapToExcel_nullType_skipped() {
        ConditionFormat cf = new ConditionFormat();
        cf.setType(null);
        cf.setCellrange(Collections.singletonList(buildRange(0, 5, 0, 0)));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    // ========== mapToExcel: DEFAULT (cell value comparison) ==========

    @Test
    void mapToExcel_default_betweenness() {
        addDefault("betweenness", Arrays.asList("10", "20"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_notBetweenness() {
        addDefault("notBetweenness", Arrays.asList("10", "20"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_greatThan() {
        addDefault("greatThan", Collections.singletonList("100"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_greaterThan_alias() {
        addDefault("greaterThan", Collections.singletonList("100"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_greatEqual() {
        addDefault("greatEqual", Collections.singletonList("50"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_lessThan() {
        addDefault("lessThan", Collections.singletonList("0"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_lessEqual() {
        addDefault("lessEqual", Collections.singletonList("5"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_equal() {
        addDefault("equal", Collections.singletonList("42"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_notEqual() {
        addDefault("notEqual", Collections.singletonList("0"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_unknownConditionName_defaultsToBetween() {
        addDefault("foo", Arrays.asList("1", "2"));
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_default_emptyConditionValue() {
        addDefault("greatThan", null);
        assertNumberFormattings(1);
    }

    // ========== mapToExcel: DATA_BAR ==========

    @Test
    void mapToExcel_dataBar_default() {
        ConditionFormat cf = baseCf(ConditionFormatType.DATA_BAR);
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_dataBar_withColor() {
        ConditionFormat cf = baseCf(ConditionFormatType.DATA_BAR);
        Map<String, Object> fmt = new HashMap<>();
        fmt.put("color", "#FF0000");
        cf.setFormat(fmt);
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    // ========== mapToExcel: COLOR_GRADATION ==========

    @Test
    void mapToExcel_colorGradation_default() {
        ConditionFormat cf = baseCf(ConditionFormatType.COLOR_GRADATION);
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_colorGradation_withColors() {
        ConditionFormat cf = baseCf(ConditionFormatType.COLOR_GRADATION);
        Map<String, Object> fmt = new HashMap<>();
        fmt.put("leastcolor", "#FFEF9C");
        fmt.put("middlecolor", "#FCFCFF");
        fmt.put("maxcolor", "#63BE7B");
        cf.setFormat(fmt);
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    // ========== mapToExcel: ICONS ==========

    @Test
    void mapToExcel_icons() {
        ConditionFormat cf = baseCf(ConditionFormatType.ICONS);
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    // ========== mapToExcel: 多个 ranges ==========

    @Test
    void mapToExcel_multipleCellranges() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        cf.setCellrange(Arrays.asList(buildRange(0, 5, 0, 0), buildRange(10, 15, 1, 2)));
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_invalidRangeShape_skipped() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        Range bad = new Range();
        bad.setRow(Collections.singletonList(0));
        bad.setColumn(Collections.singletonList(0));
        cf.setCellrange(Collections.singletonList(bad));
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    @Test
    void mapToExcel_nullRangeInList_skipped() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        cf.setCellrange(Arrays.asList(null, buildRange(0, 5, 0, 0)));
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_nullRowInRange_skipped() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        Range bad = new Range();
        bad.setRow(null);
        bad.setColumn(Arrays.asList(0, 1));
        cf.setCellrange(Collections.singletonList(bad));
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    @Test
    void mapToExcel_nullColumnInRange_skipped() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        Range bad = new Range();
        bad.setRow(Arrays.asList(0, 1));
        bad.setColumn(null);
        cf.setCellrange(Collections.singletonList(bad));
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    @Test
    void mapToExcel_emptyRowInRange_skipped() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        Range bad = new Range();
        bad.setRow(new ArrayList<>());
        bad.setColumn(Arrays.asList(0, 1));
        cf.setCellrange(Collections.singletonList(bad));
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertEquals(0, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }

    @Test
    void mapToExcel_nullFormatMap_usesDefaults() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        cf.setFormat(null);
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    @Test
    void mapToExcel_emptyFormatMap_usesDefaults() {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        cf.setFormat(new HashMap<>());
        cf.setConditionName("greatThan");
        cf.setConditionValue(Collections.singletonList("0"));
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
        assertNumberFormattings(1);
    }

    // ========== mapToLuckySheet ==========

    @Test
    void mapToLuckySheet_emptyConditionalFormatting_doesNothing() {
        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getLuckysheet_conditionformat_save());
    }

    @Test
    void mapToLuckySheet_defaultGreaterThan() {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        XSSFConditionalFormattingRule rule =
                (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(ComparisonOperator.GT, "100");
        XSSFPatternFormatting pf = rule.createPatternFormatting();
        pf.setFillBackgroundColor(new XSSFColor(new byte[]{(byte) 0xFF, 0, 0}, null));
        pf.setFillPattern(XSSFPatternFormatting.SOLID_FOREGROUND);
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 0, 0)}, rule);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(1, saves.size());
        assertEquals(ConditionFormatType.DEFAULT, saves.get(0).getType());
        assertEquals("greatThan", saves.get(0).getConditionName());
    }

    @Test
    void mapToLuckySheet_defaultBetween_andNotBetween() {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule between = scf.createConditionalFormattingRule(
                ComparisonOperator.BETWEEN, "10", "20");
        ConditionalFormattingRule notBetween = scf.createConditionalFormattingRule(
                ComparisonOperator.NOT_BETWEEN, "0", "5");
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 0, 0)}, between);
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 1, 1)}, notBetween);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(2, saves.size());
    }

    @Test
    void mapToLuckySheet_allDefaultOperators() {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        addPoiRule(scf, ComparisonOperator.LT, "0", null, 0);
        addPoiRule(scf, ComparisonOperator.GE, "5", null, 1);
        addPoiRule(scf, ComparisonOperator.LE, "10", null, 2);
        addPoiRule(scf, ComparisonOperator.EQUAL, "42", null, 3);
        addPoiRule(scf, ComparisonOperator.NOT_EQUAL, "0", null, 4);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(5, saves.size());
        assertEquals("lessThan", saves.get(0).getConditionName());
        assertEquals("greatEqual", saves.get(1).getConditionName());
        assertEquals("lessEqual", saves.get(2).getConditionName());
        assertEquals("equal", saves.get(3).getConditionName());
        assertEquals("notEqual", saves.get(4).getConditionName());
    }

    @Test
    void mapToLuckySheet_dataBar() {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        XSSFColor color = new XSSFColor(new byte[]{0x63, (byte) 0x8E, (byte) 0xC6}, null);
        XSSFConditionalFormattingRule rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(color);
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 0, 0)}, rule);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(1, saves.size());
        assertEquals(ConditionFormatType.DATA_BAR, saves.get(0).getType());
        assertEquals("dataBar", saves.get(0).getConditionName());
    }

    @Test
    void mapToLuckySheet_colorScale() {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        XSSFConditionalFormattingRule rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingColorScaleRule();
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 0, 0)}, rule);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(1, saves.size());
        assertEquals(ConditionFormatType.COLOR_GRADATION, saves.get(0).getType());
    }

    @Test
    void mapToLuckySheet_iconSet() {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        XSSFConditionalFormattingRule rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(
                org.apache.poi.ss.usermodel.IconMultiStateFormatting.IconSet.GYR_3_ARROW);
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 0, 0)}, rule);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(1, saves.size());
        assertEquals(ConditionFormatType.ICONS, saves.get(0).getType());
    }

    @Test
    void mapToLuckySheet_defaultWithNullFormulas() {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule = scf.createConditionalFormattingRule(
                ComparisonOperator.GT, (String) null);
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 0, 0)}, rule);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(1, saves.size());
    }

    @Test
    void mapToLuckySheet_existingSavesList() {
        List<ConditionFormat> existing = new ArrayList<>();
        ConditionFormat existingCf = new ConditionFormat();
        existingCf.setType(ConditionFormatType.DEFAULT);
        existing.add(existingCf);
        luckySheet.setLuckysheet_conditionformat_save(existing);

        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule = scf.createConditionalFormattingRule(
                ComparisonOperator.GT, "100");
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, 0, 0)}, rule);

        ConditionFormatMapper.mapToLuckySheet(sheet, luckySheet);

        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(2, saves.size());
    }

    // ========== Round-trip ==========

    @Test
    void roundTrip_defaultGreaterThan() {
        ConditionFormat src = new ConditionFormat();
        src.setType(ConditionFormatType.DEFAULT);
        src.setCellrange(Collections.singletonList(buildRange(0, 5, 0, 0)));
        src.setConditionName("greatThan");
        src.setConditionValue(Collections.singletonList("100"));

        ConditionFormatMapper.mapToExcel(Collections.singletonList(src), sheet);

        LuckySheet result = new LuckySheet();
        ConditionFormatMapper.mapToLuckySheet(sheet, result);

        List<ConditionFormat> saves = result.getLuckysheet_conditionformat_save();
        assertNotNull(saves);
        assertEquals(1, saves.size());
        assertEquals(ConditionFormatType.DEFAULT, saves.get(0).getType());
        assertEquals("greatThan", saves.get(0).getConditionName());
        assertTrue(saves.get(0).getConditionValue().contains("100"));
    }

    // ========== helpers ==========

    private void addDefault(String conditionName, List<Object> values) {
        ConditionFormat cf = baseCf(ConditionFormatType.DEFAULT);
        cf.setConditionName(conditionName);
        cf.setConditionValue(values);
        ConditionFormatMapper.mapToExcel(Collections.singletonList(cf), sheet);
    }

    private ConditionFormat baseCf(ConditionFormatType type) {
        ConditionFormat cf = new ConditionFormat();
        cf.setType(type);
        cf.setCellrange(Collections.singletonList(buildRange(0, 5, 0, 0)));
        return cf;
    }

    private Range buildRange(int r1, int r2, int c1, int c2) {
        Range r = new Range();
        r.setRow(Arrays.asList(r1, r2));
        r.setColumn(Arrays.asList(c1, c2));
        return r;
    }

    private void addPoiRule(SheetConditionalFormatting scf, byte op, String f1, String f2, int col) {
        ConditionalFormattingRule rule = f2 != null
                ? scf.createConditionalFormattingRule(op, f1, f2)
                : scf.createConditionalFormattingRule(op, f1);
        scf.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 5, col, col)}, rule);
    }

    private void assertNumberFormattings(int expected) {
        assertEquals(expected, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
    }
}
