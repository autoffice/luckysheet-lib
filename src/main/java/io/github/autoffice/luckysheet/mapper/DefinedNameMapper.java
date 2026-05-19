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

import io.github.autoffice.luckysheet.model.DefinedName;
import io.github.autoffice.luckysheet.model.LuckyFile;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

/**
 * Workbook 级命名范围 Luckysheet ↔ POI 双向映射器.
 */
public final class DefinedNameMapper {

    private DefinedNameMapper() {
    }

    /**
     * 从 Excel 工作簿中提取命名范围并转换为 Luckysheet 格式.
     *
     * @param workbook  源 POI 工作簿
     * @param luckyFile 目标 LuckyFile
     */
    public static void mapToLuckyFile(XSSFWorkbook workbook, LuckyFile luckyFile) {
        List<XSSFName> names = workbook.getAllNames();
        if (names == null || names.isEmpty()) {
            return;
        }
        List<DefinedName> list = new ArrayList<>();
        for (XSSFName name : names) {
            if (name.getNameName() != null && name.getNameName().startsWith("_xlnm.")) {
                // 跳过内置命名 (如 _xlnm._FilterDatabase, _xlnm.Print_Area)
                continue;
            }
            DefinedName dn = new DefinedName();
            dn.setName(name.getNameName());
            dn.setFormula(name.getRefersToFormula());
            dn.setComment(name.getComment());
            dn.setSheetIndex(name.getSheetIndex());
            dn.setHidden(name.isHidden());
            list.add(dn);
        }
        if (!list.isEmpty()) {
            luckyFile.setDefinedNames(list);
        }
    }

    /**
     * 将 Luckysheet 命名范围列表写入 Excel 工作簿.
     *
     * @param definedNames 命名范围列表
     * @param workbook     目标 POI 工作簿
     */
    public static void mapToExcel(List<DefinedName> definedNames, XSSFWorkbook workbook) {
        if (CollectionUtils.isEmpty(definedNames)) {
            return;
        }
        for (DefinedName dn : definedNames) {
            if (dn == null || dn.getName() == null || dn.getFormula() == null) {
                continue;
            }
            if (workbook.getName(dn.getName()) != null) {
                continue;
            }
            XSSFName name = workbook.createName();
            name.setNameName(dn.getName());
            name.setRefersToFormula(dn.getFormula());
            if (dn.getSheetIndex() != null && dn.getSheetIndex() >= 0) {
                name.setSheetIndex(dn.getSheetIndex());
            }
            if (dn.getComment() != null) {
                name.setComment(dn.getComment());
            }
        }
    }
}
