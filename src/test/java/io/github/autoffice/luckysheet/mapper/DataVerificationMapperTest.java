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

import io.github.autoffice.luckysheet.model.sheet.DataVerification;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class DataVerificationMapperTest {

    // --- mapToLuckySheet tests ---

    @Test
    void mapToLuckySheet_noValidations() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        LuckySheet luckySheet = new LuckySheet();

        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getDataVerification());
    }

    @Test
    void mapToLuckySheet_dropdownValidation() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(
                new String[]{"A", "B", "C"});
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        Map<String, DataVerification> map = luckySheet.getDataVerification();
        assertNotNull(map);
        DataVerification v = map.get("0_0");
        assertNotNull(v);
        assertEquals("dropdown", v.getType());
        assertEquals("A,B,C", v.getValue1());
    }

    @Test
    void mapToLuckySheet_numberBetween() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(1, 1, 2, 2);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.BETWEEN, "1", "100");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("1_2");
        assertNotNull(v);
        assertEquals("number", v.getType());
        assertEquals("bw", v.getType2());
        assertEquals("1", v.getValue1());
        assertEquals("100", v.getValue2());
    }

    @Test
    void mapToLuckySheet_numberNotBetween() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.NOT_BETWEEN, "5", "50");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("number", v.getType());
        assertEquals("nb", v.getType2());
    }

    @Test
    void mapToLuckySheet_numberEqual() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.EQUAL, "10", null);
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("number", v.getType());
        assertEquals("eq", v.getType2());
    }

    @Test
    void mapToLuckySheet_numberNotEqual() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.NOT_EQUAL, "10", null);
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("number", v.getType());
        assertEquals("ne", v.getType2());
    }

    @Test
    void mapToLuckySheet_numberGreaterThan() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.GREATER_THAN, "10", null);
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("number", v.getType());
        assertEquals("gt", v.getType2());
    }

    @Test
    void mapToLuckySheet_numberLessThan() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.LESS_THAN, "10", null);
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("number", v.getType());
        assertEquals("lt", v.getType2());
    }

    @Test
    void mapToLuckySheet_numberGreaterOrEqual() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "10", null);
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("number", v.getType());
        assertEquals("gte", v.getType2());
    }

    @Test
    void mapToLuckySheet_numberLessOrEqual() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDecimalConstraint(
                DataValidationConstraint.OperatorType.LESS_OR_EQUAL, "10", null);
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("number", v.getType());
        assertEquals("lte", v.getType2());
    }

    @Test
    void mapToLuckySheet_textLength() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createTextLengthConstraint(
                DataValidationConstraint.OperatorType.BETWEEN, "1", "50");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("text_length", v.getType());
        assertEquals("bw", v.getType2());
    }

    @Test
    void mapToLuckySheet_dateWithAfter() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDateConstraint(
                DataValidationConstraint.OperatorType.GREATER_THAN,
                "2024-01-01", null, "yyyy-MM-dd");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("date", v.getType());
        assertEquals("af", v.getType2());
    }

    @Test
    void mapToLuckySheet_dateBefore() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDateConstraint(
                DataValidationConstraint.OperatorType.LESS_THAN,
                "2024-12-31", null, "yyyy-MM-dd");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("date", v.getType());
        assertEquals("bf", v.getType2());
    }

    @Test
    void mapToLuckySheet_dateNotAfter() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDateConstraint(
                DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
                "2024-01-01", null, "yyyy-MM-dd");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("date", v.getType());
        assertEquals("naf", v.getType2());
    }

    @Test
    void mapToLuckySheet_dateNotBefore() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createDateConstraint(
                DataValidationConstraint.OperatorType.LESS_OR_EQUAL,
                "2024-12-31", null, "yyyy-MM-dd");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertEquals("date", v.getType());
        assertEquals("nbf", v.getType2());
    }

    @Test
    void mapToLuckySheet_formulaType() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createCustomConstraint("LEN(A1)>0");
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertNotNull(v);
        assertEquals("text_content", v.getType());
        assertEquals("include", v.getType2());
    }

    @Test
    void mapToLuckySheet_withProhibitInputAndHint() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 0, 0, 0);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(
                new String[]{"Yes", "No"});
        DataValidation dv = helper.createValidation(constraint, range);
        dv.setErrorStyle(DataValidation.ErrorStyle.STOP);
        dv.setShowErrorBox(true);
        dv.createErrorBox("Error", "Invalid");
        dv.setShowPromptBox(true);
        dv.createPromptBox("Hint", "Select Yes or No");
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        DataVerification v = luckySheet.getDataVerification().get("0_0");
        assertNotNull(v);
        assertTrue(v.getProhibitInput());
        assertTrue(v.getHintShow());
        assertEquals("Select Yes or No", v.getHintText());
    }

    @Test
    void mapToLuckySheet_multiCellRange() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList range = new CellRangeAddressList(0, 1, 0, 1);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(
                new String[]{"X", "Y"});
        DataValidation dv = helper.createValidation(constraint, range);
        sheet.addValidationData(dv);

        LuckySheet luckySheet = new LuckySheet();
        DataVerificationMapper.mapToLuckySheet(sheet, luckySheet);

        Map<String, DataVerification> map = luckySheet.getDataVerification();
        // 2x2 range should produce 4 entries
        assertNotNull(map.get("0_0"));
        assertNotNull(map.get("0_1"));
        assertNotNull(map.get("1_0"));
        assertNotNull(map.get("1_1"));
    }

    // --- mapToExcel tests ---

    @Test
    void mapToExcel_nullMap() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataVerificationMapper.mapToExcel(null, sheet);
        assertTrue(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_emptyMap() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        DataVerificationMapper.mapToExcel(new HashMap<>(), sheet);
        assertTrue(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_invalidKeys() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("dropdown");
        dv.setValue1("A,B");
        map.put("invalid", dv);
        map.put("_1", dv);
        map.put("1_", dv);
        map.put("abc_def", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertTrue(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_nullValue() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        map.put("0_0", null);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertTrue(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_dropdown() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("dropdown");
        dv.setValue1("Option1, Option2, Option3");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_checkbox() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("checkbox");
        dv.setValue1("");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_numberWithAllType2() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();

        // Test "nb" type2
        DataVerification dvNb = new DataVerification();
        dvNb.setType("number");
        dvNb.setType2("nb");
        dvNb.setValue1("1");
        dvNb.setValue2("100");
        map.put("0_0", dvNb);

        // Test "af" type2
        DataVerification dvAf = new DataVerification();
        dvAf.setType("number");
        dvAf.setType2("af");
        dvAf.setValue1("10");
        dvAf.setValue2("");
        map.put("1_0", dvAf);

        // Test "bf" type2
        DataVerification dvBf = new DataVerification();
        dvBf.setType("number");
        dvBf.setType2("bf");
        dvBf.setValue1("10");
        dvBf.setValue2("");
        map.put("2_0", dvBf);

        // Test "naf" type2
        DataVerification dvNaf = new DataVerification();
        dvNaf.setType("number");
        dvNaf.setType2("naf");
        dvNaf.setValue1("10");
        dvNaf.setValue2("");
        map.put("3_0", dvNaf);

        // Test "nbf" type2
        DataVerification dvNbf = new DataVerification();
        dvNbf.setType("number");
        dvNbf.setType2("nbf");
        dvNbf.setValue1("10");
        dvNbf.setValue2("");
        map.put("4_0", dvNbf);

        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_textContent_include() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("text_content");
        dv.setType2("include");
        dv.setValue1("hello");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_textContent_exclude() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("text_content");
        dv.setType2("exclude");
        dv.setValue1("bad");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_textContent_equal() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("text_content");
        dv.setType2("equal");
        dv.setValue1("exact");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_validity_phone() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("validity");
        dv.setType2("phone");
        dv.setValue1("");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_validity_card() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("validity");
        dv.setType2("card");
        dv.setValue1("");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_date() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("date");
        dv.setType2("bw");
        dv.setValue1("2024-01-01");
        dv.setValue2("2024-12-31");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_textLength() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("text_length");
        dv.setType2("bw");
        dv.setValue1("1");
        dv.setValue2("100");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_unsupportedType() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("unknown_type");
        dv.setValue1("x");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertTrue(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_withProhibitInputAndHint() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("dropdown");
        dv.setValue1("A,B,C");
        dv.setProhibitInput(true);
        dv.setHintShow(true);
        dv.setHintText("Please select");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_dateWithDateType2Values() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();

        // Test "af" (after) for date
        DataVerification dvAf = new DataVerification();
        dvAf.setType("date");
        dvAf.setType2("af");
        dvAf.setValue1("2024-01-01");
        dvAf.setValue2("");
        map.put("0_0", dvAf);

        // Test "bf" (before) for date
        DataVerification dvBf = new DataVerification();
        dvBf.setType("date");
        dvBf.setType2("bf");
        dvBf.setValue1("2024-12-31");
        dvBf.setValue2("");
        map.put("1_0", dvBf);

        // Test "naf" (not after) for date
        DataVerification dvNaf = new DataVerification();
        dvNaf.setType("date");
        dvNaf.setType2("naf");
        dvNaf.setValue1("2024-06-01");
        dvNaf.setValue2("");
        map.put("2_0", dvNaf);

        // Test "nbf" (not before) for date
        DataVerification dvNbf = new DataVerification();
        dvNbf.setType("date");
        dvNbf.setType2("nbf");
        dvNbf.setValue1("2024-06-01");
        dvNbf.setValue2("");
        map.put("3_0", dvNbf);

        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_nullType2DefaultsToBetween() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("number");
        dv.setType2(null);
        dv.setValue1("1");
        dv.setValue2("10");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }

    @Test
    void mapToExcel_unknownType2DefaultsToBetween() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        Map<String, DataVerification> map = new LinkedHashMap<>();
        DataVerification dv = new DataVerification();
        dv.setType("number");
        dv.setType2("unknown_op");
        dv.setValue1("1");
        dv.setValue2("10");
        map.put("0_0", dv);
        DataVerificationMapper.mapToExcel(map, sheet);
        assertFalse(sheet.getDataValidations().isEmpty());
    }
}
