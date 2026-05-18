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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

/**
 * 工作表保护 Luckysheet ↔ POI 双向映射器.
 */
public final class ProtectionMapper {

    private ProtectionMapper() {
    }

    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        if (ctWorksheet == null || !ctWorksheet.isSetSheetProtection()) {
            return;
        }
        CTSheetProtection protection = ctWorksheet.getSheetProtection();
        Authority a = new Authority();
        a.setSheet(protection.getSheet() ? 1 : 0);
        // POI-xmlbeans: "not" 语义字段: selectLockedCells=true => 禁止选择被锁单元格
        a.setSelectLockedCells(protection.getSelectLockedCells() ? 0 : 1);
        a.setSelectunLockedCells(protection.getSelectUnlockedCells() ? 0 : 1);
        a.setFormatCells(protection.getFormatCells() ? 0 : 1);
        a.setFormatColumns(protection.getFormatColumns() ? 0 : 1);
        a.setFormatRows(protection.getFormatRows() ? 0 : 1);
        a.setInsertColumns(protection.getInsertColumns() ? 0 : 1);
        a.setInsertRows(protection.getInsertRows() ? 0 : 1);
        a.setInsertHyperlinks(protection.getInsertHyperlinks() ? 0 : 1);
        a.setDeleteColumns(protection.getDeleteColumns() ? 0 : 1);
        a.setDeleteRows(protection.getDeleteRows() ? 0 : 1);
        a.setSort(protection.getSort() ? 0 : 1);
        a.setFilter(protection.getAutoFilter() ? 0 : 1);
        a.setUsePivotTablereports(protection.getPivotTables() ? 0 : 1);
        a.setEditObjects(protection.getObjects() ? 0 : 1);
        a.setEditScenarios(protection.getScenarios() ? 0 : 1);
        if (protection.isSetAlgorithmName()) {
            a.setAlgorithmName(protection.getAlgorithmName());
        }
        if (protection.isSetSaltValue()) {
            a.setSaltValue(java.util.Base64.getEncoder().encodeToString(protection.getSaltValue()));
        }
        if (protection.isSetSpinCount()) {
            a.setSpinCount((int) protection.getSpinCount());
        }
        if (protection.isSetHashValue()) {
            a.setHashValue(java.util.Base64.getEncoder().encodeToString(protection.getHashValue()));
        }
        luckySheet.getConfig().setAuthority(a);
    }

    public static void mapToExcel(Authority authority, XSSFSheet sheet) {
        if (authority == null || authority.getSheet() == null || authority.getSheet() != 1) {
            return;
        }
        sheet.enableLocking();
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTSheetProtection protection = ctWorksheet.isSetSheetProtection()
                ? ctWorksheet.getSheetProtection()
                : ctWorksheet.addNewSheetProtection();
        protection.setSheet(true);
        applyAllow(protection::setSelectLockedCells, authority.getSelectLockedCells());
        applyAllow(protection::setSelectUnlockedCells, authority.getSelectunLockedCells());
        applyAllow(protection::setFormatCells, authority.getFormatCells());
        applyAllow(protection::setFormatColumns, authority.getFormatColumns());
        applyAllow(protection::setFormatRows, authority.getFormatRows());
        applyAllow(protection::setInsertColumns, authority.getInsertColumns());
        applyAllow(protection::setInsertRows, authority.getInsertRows());
        applyAllow(protection::setInsertHyperlinks, authority.getInsertHyperlinks());
        applyAllow(protection::setDeleteColumns, authority.getDeleteColumns());
        applyAllow(protection::setDeleteRows, authority.getDeleteRows());
        applyAllow(protection::setSort, authority.getSort());
        applyAllow(protection::setAutoFilter, authority.getFilter());
        applyAllow(protection::setPivotTables, authority.getUsePivotTablereports());
        applyAllow(protection::setObjects, authority.getEditObjects());
        applyAllow(protection::setScenarios, authority.getEditScenarios());

        if (authority.getAlgorithmName() != null) {
            protection.setAlgorithmName(authority.getAlgorithmName());
        }
        if (authority.getSaltValue() != null) {
            try {
                protection.setSaltValue(java.util.Base64.getDecoder().decode(authority.getSaltValue()));
            } catch (IllegalArgumentException ignore) {
                // 无效 base64 盐值不致命, 丢弃
                authority.setSaltValue(null);
            }
        }
        if (authority.getSpinCount() != null) {
            protection.setSpinCount(authority.getSpinCount());
        }
        if (authority.getHashValue() != null) {
            try {
                protection.setHashValue(java.util.Base64.getDecoder().decode(authority.getHashValue()));
            } catch (IllegalArgumentException ignore) {
                // 无效 base64 哈希值丢弃
                authority.setHashValue(null);
            }
        }
    }

    private static void applyAllow(java.util.function.Consumer<Boolean> setter, Integer luckyAllowValue) {
        if (luckyAllowValue == null) {
            return;
        }
        // Luckysheet: 1 = 允许; Excel CTSheetProtection: true = 禁止
        setter.accept(luckyAllowValue != 1);
    }
}
