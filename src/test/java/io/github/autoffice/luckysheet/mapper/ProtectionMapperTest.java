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

import io.github.autoffice.luckysheet.model.sheet.Authority;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.SheetConfig;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import java.io.IOException;
import java.util.Base64;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 测试 {@link ProtectionMapper} 的双向映射逻辑, 覆盖工作表保护的各种配置场景.
 */
class ProtectionMapperTest {

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
        luckySheet.setConfig(new SheetConfig());
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    // ========== mapToLuckySheet ==========

    @Test
    void mapToLuckySheet_nullCTWorksheet_doesNothing() {
        // CTWorksheet is null by default in a fresh sheet, but let's be explicit
        ProtectionMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getConfig().getAuthority());
    }

    @Test
    void mapToLuckySheet_noSheetProtection_doesNothing() {
        // Ensure CTWorksheet exists but has no protection
        sheet.getCTWorksheet(); // Initialize it
        ProtectionMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getConfig().getAuthority());
    }

    @Test
    void mapToLuckySheet_fullProtection_allFlagsTrue() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTSheetProtection protection = ctWorksheet.addNewSheetProtection();
        protection.setSheet(true);
        protection.setSelectLockedCells(true);
        protection.setSelectUnlockedCells(true);
        protection.setFormatCells(true);
        protection.setFormatColumns(true);
        protection.setFormatRows(true);
        protection.setInsertColumns(true);
        protection.setInsertRows(true);
        protection.setInsertHyperlinks(true);
        protection.setDeleteColumns(true);
        protection.setDeleteRows(true);
        protection.setSort(true);
        protection.setAutoFilter(true);
        protection.setPivotTables(true);
        protection.setObjects(true);
        protection.setScenarios(true);

        ProtectionMapper.mapToLuckySheet(sheet, luckySheet);

        Authority a = luckySheet.getConfig().getAuthority();
        assertNotNull(a);
        assertEquals(1, a.getSheet());
        // POI true = prohibited, so Luckysheet should be 0 (not allowed)
        assertEquals(0, a.getSelectLockedCells());
        assertEquals(0, a.getSelectunLockedCells());
        assertEquals(0, a.getFormatCells());
        assertEquals(0, a.getFormatColumns());
        assertEquals(0, a.getFormatRows());
        assertEquals(0, a.getInsertColumns());
        assertEquals(0, a.getInsertRows());
        assertEquals(0, a.getInsertHyperlinks());
        assertEquals(0, a.getDeleteColumns());
        assertEquals(0, a.getDeleteRows());
        assertEquals(0, a.getSort());
        assertEquals(0, a.getFilter());
        assertEquals(0, a.getUsePivotTablereports());
        assertEquals(0, a.getEditObjects());
        assertEquals(0, a.getEditScenarios());
    }

    @Test
    void mapToLuckySheet_allFlagsFalse() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTSheetProtection protection = ctWorksheet.addNewSheetProtection();
        protection.setSheet(true);
        protection.setSelectLockedCells(false);
        protection.setSelectUnlockedCells(false);
        protection.setFormatCells(false);
        protection.setFormatColumns(false);
        protection.setFormatRows(false);
        protection.setInsertColumns(false);
        protection.setInsertRows(false);
        protection.setInsertHyperlinks(false);
        protection.setDeleteColumns(false);
        protection.setDeleteRows(false);
        protection.setSort(false);
        protection.setAutoFilter(false);
        protection.setPivotTables(false);
        protection.setObjects(false);
        protection.setScenarios(false);

        ProtectionMapper.mapToLuckySheet(sheet, luckySheet);

        Authority a = luckySheet.getConfig().getAuthority();
        assertNotNull(a);
        // POI false = allowed, so Luckysheet should be 1 (allowed)
        assertEquals(1, a.getSelectLockedCells());
        assertEquals(1, a.getSelectunLockedCells());
        assertEquals(1, a.getFormatCells());
        assertEquals(1, a.getFormatColumns());
        assertEquals(1, a.getFormatRows());
        assertEquals(1, a.getInsertColumns());
        assertEquals(1, a.getInsertRows());
        assertEquals(1, a.getInsertHyperlinks());
        assertEquals(1, a.getDeleteColumns());
        assertEquals(1, a.getDeleteRows());
        assertEquals(1, a.getSort());
        assertEquals(1, a.getFilter());
        assertEquals(1, a.getUsePivotTablereports());
        assertEquals(1, a.getEditObjects());
        assertEquals(1, a.getEditScenarios());
    }

    @Test
    void mapToLuckySheet_withPasswordFields() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTSheetProtection protection = ctWorksheet.addNewSheetProtection();
        protection.setSheet(true);
        byte[] salt = new byte[]{0x01, 0x02, 0x03};
        byte[] hash = new byte[]{0x0A, 0x0B, 0x0C};
        protection.setAlgorithmName("SHA-512");
        protection.setSaltValue(salt);
        protection.setSpinCount(100000);
        protection.setHashValue(hash);

        ProtectionMapper.mapToLuckySheet(sheet, luckySheet);

        Authority a = luckySheet.getConfig().getAuthority();
        assertNotNull(a);
        assertEquals("SHA-512", a.getAlgorithmName());
        assertEquals(Base64.getEncoder().encodeToString(salt), a.getSaltValue());
        assertEquals(100000, a.getSpinCount());
        assertEquals(Base64.getEncoder().encodeToString(hash), a.getHashValue());
    }

    @Test
    void mapToLuckySheet_withoutPasswordFields() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTSheetProtection protection = ctWorksheet.addNewSheetProtection();
        protection.setSheet(true);
        // Don't set any password fields

        ProtectionMapper.mapToLuckySheet(sheet, luckySheet);

        Authority a = luckySheet.getConfig().getAuthority();
        assertNotNull(a);
        assertNull(a.getAlgorithmName());
        assertNull(a.getSaltValue());
        assertNull(a.getSpinCount());
        assertNull(a.getHashValue());
    }

    // ========== mapToExcel ==========

    @Test
    void mapToExcel_nullAuthority_doesNothing() {
        ProtectionMapper.mapToExcel(null, sheet);
        assertFalse(sheet.getCTWorksheet().isSetSheetProtection());
    }

    @Test
    void mapToExcel_sheetFieldNull_doesNothing() {
        Authority authority = new Authority();
        authority.setSheet(null);
        ProtectionMapper.mapToExcel(authority, sheet);
        assertFalse(sheet.getCTWorksheet().isSetSheetProtection());
    }

    @Test
    void mapToExcel_sheetZero_skipped() {
        Authority authority = new Authority();
        authority.setSheet(0);
        ProtectionMapper.mapToExcel(authority, sheet);
        assertFalse(sheet.getCTWorksheet().isSetSheetProtection());
    }

    @Test
    void mapToExcel_sheetOne_applied() {
        Authority authority = new Authority();
        authority.setSheet(1);
        ProtectionMapper.mapToExcel(authority, sheet);
        assertTrue(sheet.getCTWorksheet().isSetSheetProtection());
        assertTrue(sheet.getCTWorksheet().getSheetProtection().getSheet());
    }

    @Test
    void mapToExcel_selectLockedCells_zero() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setSelectLockedCells(0);
        ProtectionMapper.mapToExcel(authority, sheet);
        // Luckysheet 0 = not allowed, POI true = prohibited
        assertTrue(sheet.getCTWorksheet().getSheetProtection().getSelectLockedCells());
    }

    @Test
    void mapToExcel_selectLockedCells_one() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setSelectLockedCells(1);
        ProtectionMapper.mapToExcel(authority, sheet);
        // Luckysheet 1 = allowed, POI false = not prohibited
        assertFalse(sheet.getCTWorksheet().getSheetProtection().getSelectLockedCells());
    }

    @Test
    void mapToExcel_allIndividualFlags_zero() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setSelectLockedCells(0);
        authority.setSelectunLockedCells(0);
        authority.setFormatCells(0);
        authority.setFormatColumns(0);
        authority.setFormatRows(0);
        authority.setInsertColumns(0);
        authority.setInsertRows(0);
        authority.setInsertHyperlinks(0);
        authority.setDeleteColumns(0);
        authority.setDeleteRows(0);
        authority.setSort(0);
        authority.setFilter(0);
        authority.setUsePivotTablereports(0);
        authority.setEditObjects(0);
        authority.setEditScenarios(0);

        ProtectionMapper.mapToExcel(authority, sheet);

        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertTrue(p.getSelectLockedCells());
        assertTrue(p.getSelectUnlockedCells());
        assertTrue(p.getFormatCells());
        assertTrue(p.getFormatColumns());
        assertTrue(p.getFormatRows());
        assertTrue(p.getInsertColumns());
        assertTrue(p.getInsertRows());
        assertTrue(p.getInsertHyperlinks());
        assertTrue(p.getDeleteColumns());
        assertTrue(p.getDeleteRows());
        assertTrue(p.getSort());
        assertTrue(p.getAutoFilter());
        assertTrue(p.getPivotTables());
        assertTrue(p.getObjects());
        assertTrue(p.getScenarios());
    }

    @Test
    void mapToExcel_allIndividualFlags_one() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setSelectLockedCells(1);
        authority.setSelectunLockedCells(1);
        authority.setFormatCells(1);
        authority.setFormatColumns(1);
        authority.setFormatRows(1);
        authority.setInsertColumns(1);
        authority.setInsertRows(1);
        authority.setInsertHyperlinks(1);
        authority.setDeleteColumns(1);
        authority.setDeleteRows(1);
        authority.setSort(1);
        authority.setFilter(1);
        authority.setUsePivotTablereports(1);
        authority.setEditObjects(1);
        authority.setEditScenarios(1);

        ProtectionMapper.mapToExcel(authority, sheet);

        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertFalse(p.getSelectLockedCells());
        assertFalse(p.getSelectUnlockedCells());
        assertFalse(p.getFormatCells());
        assertFalse(p.getFormatColumns());
        assertFalse(p.getFormatRows());
        assertFalse(p.getInsertColumns());
        assertFalse(p.getInsertRows());
        assertFalse(p.getInsertHyperlinks());
        assertFalse(p.getDeleteColumns());
        assertFalse(p.getDeleteRows());
        assertFalse(p.getSort());
        assertFalse(p.getAutoFilter());
        assertFalse(p.getPivotTables());
        assertFalse(p.getObjects());
        assertFalse(p.getScenarios());
    }

    @Test
    void mapToExcel_validBase64SaltValue() {
        Authority authority = new Authority();
        authority.setSheet(1);
        String validBase64 = Base64.getEncoder().encodeToString(new byte[]{0x01, 0x02, 0x03});
        authority.setSaltValue(validBase64);

        ProtectionMapper.mapToExcel(authority, sheet);

        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertNotNull(p.getSaltValue());
        assertArrayEquals(new byte[]{0x01, 0x02, 0x03}, p.getSaltValue());
    }

    @Test
    void mapToExcel_invalidBase64SaltValue_caught() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setSaltValue("not-valid-base64!!!");

        ProtectionMapper.mapToExcel(authority, sheet);

        // Should not throw, and saltValue should be cleared
        assertNull(authority.getSaltValue());
        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertFalse(p.isSetSaltValue());
    }

    @Test
    void mapToExcel_validBase64HashValue() {
        Authority authority = new Authority();
        authority.setSheet(1);
        String validBase64 = Base64.getEncoder().encodeToString(new byte[]{0x0A, 0x0B, 0x0C});
        authority.setHashValue(validBase64);

        ProtectionMapper.mapToExcel(authority, sheet);

        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertNotNull(p.getHashValue());
        assertArrayEquals(new byte[]{0x0A, 0x0B, 0x0C}, p.getHashValue());
    }

    @Test
    void mapToExcel_invalidBase64HashValue_caught() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setHashValue("invalid-hash@#$");

        ProtectionMapper.mapToExcel(authority, sheet);

        // Should not throw, and hashValue should be cleared
        assertNull(authority.getHashValue());
        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertFalse(p.isSetHashValue());
    }

    @Test
    void mapToExcel_withSpinCount() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setSpinCount(50000);

        ProtectionMapper.mapToExcel(authority, sheet);

        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertEquals(50000, p.getSpinCount());
    }

    @Test
    void mapToExcel_withAlgorithmName() {
        Authority authority = new Authority();
        authority.setSheet(1);
        authority.setAlgorithmName("SHA-256");

        ProtectionMapper.mapToExcel(authority, sheet);

        CTSheetProtection p = sheet.getCTWorksheet().getSheetProtection();
        assertEquals("SHA-256", p.getAlgorithmName());
    }

    // ========== Round-trip ==========

    @Test
    void roundTrip_fullProtection() {
        Authority src = new Authority();
        src.setSheet(1);
        src.setSelectLockedCells(0);
        src.setSelectunLockedCells(1);
        src.setFormatCells(0);
        src.setAlgorithmName("SHA-512");
        src.setSaltValue(Base64.getEncoder().encodeToString(new byte[]{0x01, 0x02}));
        src.setSpinCount(100000);
        src.setHashValue(Base64.getEncoder().encodeToString(new byte[]{0x0A, 0x0B}));

        ProtectionMapper.mapToExcel(src, sheet);

        LuckySheet result = new LuckySheet();
        result.setConfig(new SheetConfig());
        ProtectionMapper.mapToLuckySheet(sheet, result);

        Authority dest = result.getConfig().getAuthority();
        assertNotNull(dest);
        assertEquals(1, dest.getSheet());
        assertEquals(0, dest.getSelectLockedCells());
        assertEquals(1, dest.getSelectunLockedCells());
        assertEquals(0, dest.getFormatCells());
        assertEquals("SHA-512", dest.getAlgorithmName());
        assertEquals(src.getSaltValue(), dest.getSaltValue());
        assertEquals(100000, dest.getSpinCount());
        assertEquals(src.getHashValue(), dest.getHashValue());
    }
}
